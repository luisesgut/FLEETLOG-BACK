"""
FleetLog Backend
----------------
Gateway entre el frontend y la API GPS de SoftwareGM.

- Hace polling al endpoint seguimientoUnidades (el token NUNCA sale del servidor)
- Persiste metadata del admin (conductor, visibilidad, load, estado) en SQLite
- Sincroniza Excel de SharePoint/Local a Supabase y mantiene caché en memoria
- Mergea GPS + metadata + datos de Excel y lo expone vía REST y WebSocket

Ejecutar:
    uvicorn main:app --reload
"""

import asyncio
import base64

import msal
import io
import json
import logging
import os
import sqlite3
import threading
import uuid
from contextlib import asynccontextmanager
from datetime import datetime, timezone

import httpx
import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from supabase import Client, create_client


load_dotenv()

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
GPS_API_URL = os.getenv(
    "GPS_API_URL",
    "https://localizacion.softwaregm.com.mx/gps/api/cliente/equipos/seguimientoUnidades",
)
GPS_AUTH_TOKEN = os.getenv("GPS_AUTH_TOKEN", "")
GPS_CLIENTE_ID = os.getenv("GPS_CLIENTE_ID", "")
POLL_INTERVAL = int(os.getenv("POLL_INTERVAL_SECONDS", "35"))
DB_PATH = os.getenv("DB_PATH", "fleet_meta.db")
EXCEL_LOCAL_PATH = os.getenv("EXCEL_LOCAL_PATH", "/Users/luisespinosa/Downloads/2026 REPORTE OPERACIONES.xlsm")

GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_TENANT_ID = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")
EXCEL_SHARE_URL = os.getenv("EXCEL_SHARE_URL", "https://tibioflex.sharepoint.com/:x:/s/COMERCIO/IQDaareP_tHtRJEJmQETxk2zAciIYjkYFKcasykKIvZX3YE?e=SXYrlg")






_msal_app = msal.ConfidentialClientApplication(
    GRAPH_CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
    client_credential=GRAPH_CLIENT_SECRET,
)
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
EXCEL_SYNC_INTERVAL = int(os.getenv("EXCEL_SYNC_INTERVAL_SECONDS", "300"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("fleetlog")

VALID_TRUCK_STATUS = {"disponible", "en_ruta", "mantenimiento", "fuera_servicio"}
VALID_LOAD_STATUS = {"pendiente", "en_transito", "entregado", "cancelado"}

# Inicializar cliente de Supabase
supabase_client: Client | None = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)

# Caché en memoria para embarques del Excel
embarques_excel_cache: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# Persistencia: metadata del admin en SQLite
# ---------------------------------------------------------------------------
_db_lock = threading.Lock()


def _db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with _db_lock, _db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS truck_meta (
                id            TEXT PRIMARY KEY,
                driver        TEXT NOT NULL DEFAULT '',
                status        TEXT,            -- override manual; NULL = derivar del GPS
                visible       INTEGER NOT NULL DEFAULT 0,
                load_json     TEXT             -- load asignado (JSON) o NULL
            )
            """
        )
def _graph_token() -> str:
    r = _msal_app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
    if not r:
        r = _msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        raise RuntimeError(f"Token Graph falló: {r.get('error_description', r)}")
    return r["access_token"]

async def _descargar_excel_bytes() -> bytes:
    token = _graph_token()
    sid = "u!" + base64.urlsafe_b64encode(EXCEL_SHARE_URL.encode()).decode().rstrip("=")
    async with httpx.AsyncClient(timeout=30) as client:
        res = await client.get(
            f"https://graph.microsoft.com/v1.0/shares/{sid}/driveItem/content",
            headers={"Authorization": f"Bearer {token}"},
            follow_redirects=True,
        )
        res.raise_for_status()
        return res.content

def normalizar_clave(texto: str) -> str:

    s = str(texto or "")
    s = s.split("/")[0]                    # corta en el primer '/'
    return s.replace(" ", "").replace("-", "").strip().upper()
# --- FUNCIÓN DE SINCRONIZACIÓN EXCEL -> SUPABASE ---
async def sync_excel_to_supabase() -> None:
    """Lee la pestaña EXPO del Excel, filtra STATUS != 7 e ignora registros ECO."""
    global embarques_excel_cache

    try:
        logger.info("Iniciando procesamiento del Excel...")

        # 1. Obtener el contenido (Archivo local o SharePoint URL)
        # 1. Obtener el contenido
        if EXCEL_LOCAL_PATH and os.path.exists(EXCEL_LOCAL_PATH):
            logger.info("Cargando Excel desde archivo local: %s", EXCEL_LOCAL_PATH)
            excel_source = EXCEL_LOCAL_PATH
        else:
            logger.info("Descargando Excel desde SharePoint vía Graph...")
            excel_source = io.BytesIO(await _descargar_excel_bytes())

        # 2. Leer específicamente la pestaña 'EXPO'
        df = pd.read_excel(excel_source, sheet_name="EXPO", engine="openpyxl")
        df = df.fillna("")

        registros_dict = {}

        for _, row in df.iterrows():
            caja_raw = str(row.get("No. Caja / Unidad", "")).strip()
            status_val = str(row.get("STATUS", "")).strip()

            caja_norm = normalizar_clave(caja_raw)
            # REGLA 1: Ignorar vacíos o unidades que empiecen con 'ECO'
            if not caja_norm or caja_norm.lower() == "nan" or caja_norm.startswith("ECO"):
                continue

            # REGLA 2: Ignorar registros con STATUS 7 (ej. "7. Cerrado" o que empiece con 7)
            if status_val.startswith("7"):
                continue

            item = {
                "numero_caja": caja_norm,  # Guardamos clave normalizada
                "cliente": str(row.get("Cliente", "")).strip(),
                "destino": str(row.get("Destino", "")).strip(),
                "carrier": str(row.get("Carrier", "")).strip(),
                "aduana": str(row.get("Aduana", "")).strip(),
                "no_factura": str(row.get("No. Factura", "")).strip(),
                "caja_cruce": str(row.get("Caja de cruce / Unidad", "")).strip(),
                "status": status_val,
            }

            # Deduplica asignando por clave limpia sin espacios
            registros_dict[caja_norm] = item

        registros = list(registros_dict.values())
        nuevo_cache = registros_dict

        # 3. Guardar en Supabase (si está configurado) y actualizar caché
        if supabase_client and registros:
            supabase_client.table("excel_embarques").upsert(
                registros, on_conflict="numero_caja"
            ).execute()

            logger.info(
                "Excel sincronizado: %d registros insertados/actualizados en Supabase.",
                len(registros),
            )

        embarques_excel_cache = nuevo_cache
        logger.info("Caché local actualizado con %d embarques activos.", len(nuevo_cache))

    except Exception as e:
        logger.exception("Error durante la sincronización del Excel: %s", e)


# --- LOOP EN SEGUNDO PLANO PARA EXCEL ---
async def excel_sync_loop() -> None:
    while True:
        await sync_excel_to_supabase()
        await asyncio.sleep(EXCEL_SYNC_INTERVAL)


def get_all_meta() -> dict[str, dict]:
    with _db_lock, _db() as conn:
        rows = conn.execute("SELECT * FROM truck_meta").fetchall()
    result = {}
    for r in rows:
        result[r["id"]] = {
            "driver": r["driver"],
            "status": r["status"],
            "visible": bool(r["visible"]),
            "load": json.loads(r["load_json"]) if r["load_json"] else None,
        }
    return result


def upsert_meta(truck_id: str, **fields) -> None:
    """Crea la fila si no existe y actualiza solo los campos provistos."""
    with _db_lock, _db() as conn:
        conn.execute("INSERT OR IGNORE INTO truck_meta (id) VALUES (?)", (truck_id,))
        for key, value in fields.items():
            if key == "load":
                conn.execute(
                    "UPDATE truck_meta SET load_json = ? WHERE id = ?",
                    (json.dumps(value) if value is not None else None, truck_id),
                )
            elif key in ("driver", "status"):
                conn.execute(
                    f"UPDATE truck_meta SET {key} = ? WHERE id = ?",
                    (value, truck_id),
                )
            elif key == "visible":
                conn.execute(
                    "UPDATE truck_meta SET visible = ? WHERE id = ?",
                    (1 if value else 0, truck_id),
                )


# ---------------------------------------------------------------------------
# Estado en memoria (último snapshot GPS crudo, ya normalizado)
# ---------------------------------------------------------------------------
class AppState:
    def __init__(self) -> None:
        self.unidades_gps: list[dict] = []
        self.last_update: str | None = None
        self.last_error: str | None = None
        self.poll_count: int = 0


state = AppState()


# ---------------------------------------------------------------------------
# WebSocket manager
# ---------------------------------------------------------------------------
class ConnectionManager:
    def __init__(self) -> None:
        self.active: list[WebSocket] = []
        self._lock = asyncio.Lock()

    async def connect(self, ws: WebSocket) -> None:
        await ws.accept()
        async with self._lock:
            self.active.append(ws)

    async def disconnect(self, ws: WebSocket) -> None:
        async with self._lock:
            if ws in self.active:
                self.active.remove(ws)

    async def broadcast(self, message: dict) -> None:
        async with self._lock:
            dead = []
            for ws in self.active:
                try:
                    await ws.send_json(message)
                except Exception:
                    dead.append(ws)
            for ws in dead:
                self.active.remove(ws)


manager = ConnectionManager()


# ---------------------------------------------------------------------------
# Normalización GPS
# ---------------------------------------------------------------------------
def normalizar_equipo(eq: dict) -> dict:
    ubic = eq.get("ubicacionActual") or {}
    coords = (ubic.get("ubic") or {}).get("coordinates") or [None, None]
    cliente = eq.get("clienteAsignado") or {}
    tipo = eq.get("tipoEquipo") or {}

    # Aseguramos que la fecha lleve 'Z' al final para que el navegador la lea como UTC
    fh_raw = str(ubic.get("fh") or "").strip()
    if fh_raw and not fh_raw.endswith("Z") and not "+" in fh_raw:
        fh_utc = f"{fh_raw}Z"
    else:
        fh_utc = fh_raw

    return {
        "id": eq.get("_id"),
        "unidad": cliente.get("unidad"),
        "placas": cliente.get("placas"),
        "producto": cliente.get("producto"),
        "equipo": {
            "tipo": tipo.get("tipoEquipo"),
            "modelo": tipo.get("modelo"),
            "imei": eq.get("imei"),
            "sim": eq.get("noTelefono"),
            "proveedor": eq.get("proveedor"),
        },
        "posicion": {"lat": coords[1], "lng": coords[0]},
        "velocidad": ubic.get("vel"),
        "heading": ubic.get("hea"),
        "evento": ubic.get("evento"),
        "conectado": ubic.get("conectado"),
        "direccion": ubic.get("direccion"),
        "fechaHoraGps": fh_utc,  # <--- Enviamos la fecha estándar ISO con 'Z'
        "bateriaInternaPct": eq.get("voltBatIntP"),
        "estatusGps": eq.get("estatusDescripcion"),
    }


def merge_con_meta(unidades: list[dict]) -> list[dict]:
    """Combina snapshot GPS + metadata SQLite + datos enriquecidos de Excel/Supabase."""
    meta = get_all_meta()
    merged = []

    for u in unidades:
        num_unidad_norm = normalizar_clave(u.get("unidad"))

        # FILTRO EXPLICITO: Si la unidad de la API GPS empieza con ECO, se excluye completamente del resultado
        if num_unidad_norm.startswith("ECO"):
            continue

        m = meta.get(u["id"], {})
        datos_excel = embarques_excel_cache.get(num_unidad_norm, {})

        merged.append(
            {
                **u,
                "driver": m.get("driver", ""),
                "statusOverride": m.get("status"),
                "visible": m.get("visible", False),
                "load": m.get("load"),
                "embarqueExcel": {
                    "cliente": datos_excel.get("cliente", ""),
                    "destino": datos_excel.get("destino", ""),
                    "carrier": datos_excel.get("carrier", ""),
                    "aduana": datos_excel.get("aduana", ""),
                    "noFactura": datos_excel.get("no_factura", ""),
                    "cajaCruce": datos_excel.get("caja_cruce", ""),
                    "status": datos_excel.get("status", ""),
                }
                if datos_excel
                else None,
            }
        )
    return merged


async def broadcast_snapshot() -> None:
    await manager.broadcast(
        {
            "type": "unidades_update",
            "timestamp": state.last_update,
            "unidades": merge_con_meta(state.unidades_gps),
        }
    )


# ---------------------------------------------------------------------------
# Polling GPS
# ---------------------------------------------------------------------------
async def fetch_unidades(client: httpx.AsyncClient) -> list[dict] | None:
    headers = {"Authorization": f"Bearer {GPS_AUTH_TOKEN}", "Content-Type": "application/json"}
    body = {
        "cliente_id": GPS_CLIENTE_ID,
        "frecuenciaActualizacion": 0,
        "tiempoActualizacion": False,
    }
    try:
        r = await client.post(GPS_API_URL, json=body, headers=headers)
        r.raise_for_status()
        data = r.json()
        if data.get("ok") != 1:
            state.last_error = f"Respuesta con ok={data.get('ok')}: {r.text[:300]}"
            logger.warning(state.last_error)
            return None
        return [normalizar_equipo(eq) for eq in data.get("equipos", [])]
    except httpx.HTTPError as e:
        state.last_error = f"Error HTTP: {e}"
        logger.error(state.last_error)
    except Exception as e:
        state.last_error = f"Error inesperado: {e}"
        logger.exception("Error en fetch")
    return None


def _hubo_cambios(anterior: list[dict], nuevo: list[dict]) -> bool:
    if not anterior:
        return True
    prev = {u["id"]: u.get("fechaHoraGps") for u in anterior}
    if len(anterior) != len(nuevo):
        return True
    return any(prev.get(u["id"]) != u.get("fechaHoraGps") for u in nuevo)


async def polling_loop() -> None:
    logger.info("Polling cada %d segundos", POLL_INTERVAL)
    async with httpx.AsyncClient(timeout=20) as client:
        while True:
            unidades = await fetch_unidades(client)
            if unidades is not None:
                cambio = _hubo_cambios(state.unidades_gps, unidades)
                state.unidades_gps = unidades
                state.last_update = datetime.now(timezone.utc).isoformat()
                state.last_error = None
                state.poll_count += 1
                if cambio:
                    await broadcast_snapshot()
                    logger.info("Poll #%d: %d unidades (broadcast)", state.poll_count, len(unidades))
            await asyncio.sleep(POLL_INTERVAL)


# ---------------------------------------------------------------------------
# App Lifespan
# ---------------------------------------------------------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    # Tarea 1: Polling GPS
    task_gps = asyncio.create_task(polling_loop())
    # Tarea 2: Sync Excel
    task_excel = asyncio.create_task(excel_sync_loop())
    yield
    task_gps.cancel()
    task_excel.cancel()
    try:
        await task_gps
        await task_excel
    except asyncio.CancelledError:
        pass


app = FastAPI(title="FleetLog Backend", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------- Modelos de request ----------------------------
class MetaPatch(BaseModel):
    driver: str | None = None
    status: str | None = None
    visible: bool | None = None


class LoadInput(BaseModel):
    shipmentNumber: str
    origin: str
    destination: str
    status: str = "pendiente"


class LoadStatusPatch(BaseModel):
    status: str


# ------------------------------- Endpoints ---------------------------------
@app.get("/api/unidades")
async def get_unidades():
    return {
        "unidades": merge_con_meta(state.unidades_gps),
        "lastUpdate": state.last_update,
        "lastError": state.last_error,
    }



@app.patch("/api/unidades/{truck_id}/meta")
async def patch_meta(truck_id: str, body: MetaPatch):
    fields: dict = {}
    if body.driver is not None:
        fields["driver"] = body.driver.strip()
    if body.status is not None:
        if body.status == "":
            fields["status"] = None
        elif body.status in VALID_TRUCK_STATUS:
            fields["status"] = body.status
        else:
            raise HTTPException(400, f"Estado inválido: {body.status}")
    if body.visible is not None:
        fields["visible"] = body.visible
    if not fields:
        raise HTTPException(400, "Nada que actualizar")

    upsert_meta(truck_id, **fields)
    await broadcast_snapshot()
    return {"ok": True}


@app.put("/api/unidades/{truck_id}/load")
async def put_load(truck_id: str, body: LoadInput):
    if body.status not in VALID_LOAD_STATUS:
        raise HTTPException(400, f"Estado de load inválido: {body.status}")
    load = {
        "id": str(uuid.uuid4()),
        "shipmentNumber": body.shipmentNumber.strip().upper(),
        "origin": body.origin.strip(),
        "destination": body.destination.strip(),
        "status": body.status,
    }
    upsert_meta(truck_id, load=load)
    await broadcast_snapshot()
    return load


@app.patch("/api/unidades/{truck_id}/load")
async def patch_load(truck_id: str, body: LoadStatusPatch):
    if body.status not in VALID_LOAD_STATUS:
        raise HTTPException(400, f"Estado de load inválido: {body.status}")
    meta = get_all_meta().get(truck_id)
    if not meta or not meta.get("load"):
        raise HTTPException(404, "La unidad no tiene load asignado")
    load = {**meta["load"], "status": body.status}
    upsert_meta(truck_id, load=load)
    await broadcast_snapshot()
    return load


@app.delete("/api/unidades/{truck_id}/load")
async def delete_load(truck_id: str):
    upsert_meta(truck_id, load=None)
    await broadcast_snapshot()
    return {"ok": True}


@app.get("/api/health")
async def health():
    return {
        "status": "ok" if state.last_error is None else "degraded",
        "pollCount": state.poll_count,
        "lastUpdate": state.last_update,
        "lastError": state.last_error,
        "dashboardsConectados": len(manager.active),
    }


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await manager.connect(ws)
    try:
        await ws.send_json(
            {
                "type": "snapshot",
                "timestamp": state.last_update,
                "unidades": merge_con_meta(state.unidades_gps),
            }
        )
        while True:
            await ws.receive_text()
    except WebSocketDisconnect:
        await manager.disconnect(ws)
    except Exception:
        await manager.disconnect(ws)