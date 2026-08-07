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
from typing import Optional, Dict

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
from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from supabase import Client, create_client
from auth import verificar_token, get_perfil_usuario
from auth import verificar_token_ws

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



class RegistroInput(BaseModel):
    email: str
    password: str
    codigo: str

class OverrideInput(BaseModel):
        destino: Optional[str] = None
        carrier: Optional[str] = None
        aduana: Optional[str] = None
        noFactura: Optional[str] = None


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
overrides_cache: Dict[str, dict] = {}
# Vista admin: TODAS las filas abiertas del Excel (sin dedup por caja) + sus overrides
embarques_filas_cache: list[dict] = []
fila_overrides_cache: Dict[str, dict] = {}


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

async def cargar_overrides():
        """Lee los overrides de Supabase al cache. Llamar al arrancar y tras cada edición."""
        global overrides_cache
        try:
            res = supabase_client.table("embarque_overrides").select("*").execute()
            overrides_cache = {r["numero_caja"]: r for r in (res.data or [])}
            logger.info("Overrides cargados: %d", len(overrides_cache))
        except Exception as e:
            logger.error("Error cargando overrides: %s", e)

async def cargar_fila_overrides():
        """Lee los overrides POR FILA (tracking url, comentario, load/orden) de Supabase."""
        global fila_overrides_cache
        try:
            res = supabase_client.table("embarque_fila_overrides").select("*").execute()
            fila_overrides_cache = {r["fila_key"]: r for r in (res.data or [])}
            logger.info("Fila-overrides cargados: %d", len(fila_overrides_cache))
        except Exception as e:
            logger.error("Error cargando fila-overrides: %s", e)

def aplicar_override(caja: str, embarque: dict) -> dict:
        """El Excel manda; el override SOLO tapa los campos que el Excel dejó vacíos."""
        ov = overrides_cache.get(caja)
        if not ov:
            return embarque
        resultado = dict(embarque)
        mapeo = {"destino": "destino", "carrier": "carrier", "aduana": "aduana", "noFactura": "no_factura"}
        for campo_emb, campo_ov in mapeo.items():
            if not resultado.get(campo_emb) and ov.get(campo_ov):
                resultado[campo_emb] = ov[campo_ov]
                resultado[f"{campo_emb}Manual"] = True  # para que el front lo pinte distinto
        return resultado

def normalizar_cliente(nombre: str) -> str:
    """'DESTINY (084)' -> 'DESTINY'. Quita el sufijo entre paréntesis."""
    return str(nombre or "").split("(")[0].strip().upper()

def filtrar_por_cliente(unidades: list[dict], cliente: str, rol: str = "asesor") -> list[dict]:
    """ADMIN ve todo. Asesor ve todo lo de su cliente. Cliente solo lo visible."""
    if cliente == "ADMIN":
        return unidades
    cliente_norm = cliente.strip().upper()
    filtradas = []
    for u in unidades:
        emb = u.get("embarqueExcel")
        if not emb or normalizar_cliente(emb.get("cliente")) != cliente_norm:
            continue
        # El cliente final solo ve las unidades marcadas visibles
        if rol == "cliente" and not u.get("visible"):
            continue
        filtradas.append(u)
    return filtradas

def fila_key(caja: str, cliente: str, factura: str) -> str:
    """Clave única por fila: caja|cliente|factura. Distingue el caso de misma caja
    con distintos clientes (ej. TBIN205 en Quality y Bunzl)."""
    c = normalizar_clave(caja)
    cl = normalizar_cliente(cliente)
    f = str(factura or "").strip().upper()
    return f"{c}|{cl}|{f}"


def normalizar_clave(texto: str) -> str:

    s = str(texto or "")
    s = s.split("/")[0]                    # corta en el primer '/'
    return s.replace(" ", "").replace("-", "").strip().upper()
# --- FUNCIÓN DE SINCRONIZACIÓN EXCEL -> SUPABASE ---
async def sync_excel_to_supabase() -> None:
    """Lee EXPO. Arma 2 vistas:
    - por caja (vista GPS): ignora STATUS 7, dedup por caja.
    - filas admin: TODAS las filas abiertas (no empieza con 6 ni 7), sin dedup."""
    global embarques_excel_cache, embarques_filas_cache

    try:
        logger.info("Iniciando procesamiento del Excel...")

        # 1. Obtener el contenido (Archivo local o SharePoint URL)
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
        filas_admin: list[dict] = []

        for _, row in df.iterrows():
            caja_raw = str(row.get("No. Caja / Unidad", "")).strip()
            status_val = str(row.get("STATUS", "")).strip()
            cliente_raw = str(row.get("Cliente", "")).strip()
            factura_raw = str(row.get("No. Factura", "")).strip()
            carrier_raw = str(row.get("Carrier", "")).strip()

            caja_norm = normalizar_clave(caja_raw)
            # REGLA 1: Ignorar vacíos o unidades que empiecen con 'ECO' (aplica a AMBAS vistas)
            if not caja_norm or caja_norm.lower() == "nan" or caja_norm.startswith("ECO"):
                continue

            # --- VISTA ADMIN: todas las filas ABIERTAS (no empieza con 6 ni 7; blanco = abierto) ---
            es_cerrado = status_val.startswith("6") or status_val.startswith("7")
            if not es_cerrado:
                filas_admin.append({
                    "filaKey": fila_key(caja_raw, cliente_raw, factura_raw),
                    "caja": caja_raw,
                    "cliente": cliente_raw,
                    "carrier": carrier_raw,
                    "aduana": str(row.get("Aduana", "")).strip(),
                    "destino": str(row.get("Destino", "")).strip(),
                    "noFactura": factura_raw,
                    "pedimento": str(row.get("Pedimento", "")).strip(),
                    "status": status_val,
                    "esEDR": carrier_raw.upper() == "EDR",
                })

            # REGLA 2 (VISTA GPS): Ignorar registros con STATUS 7
            if status_val.startswith("7"):
                continue

            item = {
                "numero_caja": caja_norm,  # Guardamos clave normalizada
                "cliente": cliente_raw,
                "destino": str(row.get("Destino", "")).strip(),
                "carrier": carrier_raw,
                "aduana": str(row.get("Aduana", "")).strip(),
                "no_factura": factura_raw,
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
        embarques_filas_cache = filas_admin
        logger.info(
            "Caché actualizado: %d por-caja, %d filas-admin abiertas.",
            len(nuevo_cache), len(filas_admin),
        )

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
        # guardamos (websocket, cliente, rol)
        self.active: list[tuple[WebSocket, str, str]] = []
        self._lock = asyncio.Lock()

    async def connect(self, ws: WebSocket, cliente: str, rol: str) -> None:
        await ws.accept()
        async with self._lock:
            self.active.append((ws, cliente, rol))

    async def disconnect(self, ws: WebSocket) -> None:
        async with self._lock:
            self.active = [(w, c, r) for (w, c, r) in self.active if w is not ws]

    async def broadcast(self, unidades_completas: list[dict]) -> None:
        """A cada conexión le manda SOLO lo que le toca según cliente y rol."""
        async with self._lock:
            dead = []
            for ws, cliente, rol in self.active:
                try:
                    visibles = filtrar_por_cliente(unidades_completas, cliente, rol)
                    await ws.send_json({
                        "type": "unidades_update",
                        "timestamp": state.last_update,
                        "unidades": visibles,
                    })
                except Exception:
                    dead.append(ws)
            self.active = [(w, c, r) for (w, c, r) in self.active if w not in dead]
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
                "embarqueExcel": aplicar_override(
                    num_unidad_norm,
                    {
                        "cliente": datos_excel.get("cliente", ""),
                        "destino": datos_excel.get("destino", ""),
                        "carrier": datos_excel.get("carrier", ""),
                        "aduana": datos_excel.get("aduana", ""),
                        "noFactura": datos_excel.get("no_factura", ""),
                        "cajaCruce": datos_excel.get("caja_cruce", ""),
                        "status": datos_excel.get("status", ""),
                    },
                )
                if datos_excel
                else None,
            }
        )
    return merged


async def broadcast_snapshot() -> None:
    todas = merge_con_meta(state.unidades_gps)
    await manager.broadcast(todas)

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
    # Cargar overrides desde Supabase al arrancar
    if supabase_client:
        await cargar_overrides()
        await cargar_fila_overrides()
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

origins_env = os.getenv("CORS_ORIGINS", "http://localhost:3000,http://127.0.0.1:3000,https://fleetlog-front.vercel.app")
ALLOWED_ORIGINS = [o.strip() for o in origins_env.split(",") if o.strip()]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
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


class FilaOverrideInput(BaseModel):
    trackingUrl: Optional[str] = None
    comentario: Optional[str] = None
    loadOrden: Optional[str] = None


# ------------------------------- Endpoints ---------------------------------
@app.get("/api/unidades")
async def get_unidades(claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    todas = merge_con_meta(state.unidades_gps)
    visibles = filtrar_por_cliente(todas, perfil["cliente"], perfil["rol"])
    return {
        "unidades": visibles,
        "cliente": perfil["cliente"],
        "rol": perfil["rol"],
        "lastUpdate": state.last_update,
    }


@app.patch("/api/unidades/{truck_id}/meta")
async def patch_meta(truck_id: str, body: MetaPatch, claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["rol"] == "cliente":
        raise HTTPException(403, "Los clientes no pueden editar")
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
async def put_load(truck_id: str, body: LoadInput, claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["rol"] == "cliente":
        raise HTTPException(403, "Los clientes no pueden editar")
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
async def patch_load(truck_id: str, body: LoadStatusPatch, claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["rol"] == "cliente":
        raise HTTPException(403, "Los clientes no pueden editar")
    if body.status not in VALID_LOAD_STATUS:
        raise HTTPException(400, f"Estado de load inválido: {body.status}")
    meta = get_all_meta().get(truck_id)
    if not meta or not meta.get("load"):
        raise HTTPException(404, "La unidad no tiene load asignado")
    load = {**meta["load"], "status": body.status}
    upsert_meta(truck_id, load=load)
    await broadcast_snapshot()
    return load

@app.patch("/api/embarques/{numero_caja}/override")
async def patch_override(numero_caja: str, body: OverrideInput, claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["rol"] == "cliente":
        raise HTTPException(403, "Los clientes no pueden editar")

    caja = normalizar_clave(numero_caja)

    # Verificar que el embarque exista y sea del cliente del asesor (o admin)
    emb = embarques_excel_cache.get(caja)
    if not emb:
        raise HTTPException(404, "Embarque no encontrado")
    if perfil["cliente"] != "ADMIN" and normalizar_cliente(emb.get("cliente")) != perfil["cliente"]:
        raise HTTPException(403, "No puedes editar embarques de otro cliente")

    datos = {
        "numero_caja": caja,
        "editado_por": claims.get("sub"),
        "editado_en": datetime.now(timezone.utc).isoformat(),
    }
    if body.destino is not None: datos["destino"] = body.destino.strip()
    if body.carrier is not None: datos["carrier"] = body.carrier.strip()
    if body.aduana is not None: datos["aduana"] = body.aduana.strip()
    if body.noFactura is not None: datos["no_factura"] = body.noFactura.strip()

    supabase_client.table("embarque_overrides").upsert(datos, on_conflict="numero_caja").execute()

    # Refrescar el cache de overrides y avisar por WS
    await cargar_overrides()
    await broadcast_snapshot()
    return {"ok": True, "numero_caja": caja}


@app.delete("/api/unidades/{truck_id}/load")
async def delete_load(truck_id: str, claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["rol"] == "cliente":
        raise HTTPException(403, "Los clientes no pueden editar")
    upsert_meta(truck_id, load=None)
    await broadcast_snapshot()
    return {"ok": True}


# ---------------------------------------------------------------------------
# Vista ADMIN de embarques: seccionada en 3 grupos
#   - edr:          filas del Excel con carrier EDR (con GPS pegado si reporta)
#   - gpsSinExcel:  unidades reportando en GPS SIN fila abierta en el Excel
#   - otrosCarriers: filas del Excel de carriers no-EDR (necesitan liga manual)
# ---------------------------------------------------------------------------
@app.get("/api/embarques")
async def get_embarques(claims: dict = Depends(verificar_token)):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["cliente"] != "ADMIN":
        raise HTTPException(403, "Solo el administrador puede ver esta vista")

    # GPS por caja normalizada
    unidades_gps = merge_con_meta(state.unidades_gps)
    gps_por_caja = {}
    for u in unidades_gps:
        gps_por_caja[normalizar_clave(u.get("unidad"))] = u

    def _armar_fila(f: dict, gps: dict | None) -> dict:
        ov = fila_overrides_cache.get(f["filaKey"], {})
        return {
            **f,
            "trackingUrl": ov.get("tracking_url", ""),
            "comentario": ov.get("comentario", ""),
            "loadOrden": ov.get("load_orden", ""),
            "gps": {
                "posicion": gps.get("posicion"),
                "velocidad": gps.get("velocidad"),
                "fechaHoraGps": gps.get("fechaHoraGps"),
                "direccion": gps.get("direccion"),
            } if gps else None,
        }

    edr: list[dict] = []
    otros_carriers: list[dict] = []
    cajas_excel_edr: set[str] = set()

    for f in embarques_filas_cache:
        if f["esEDR"]:
            caja_n = normalizar_clave(f["caja"])
            cajas_excel_edr.add(caja_n)
            edr.append(_armar_fila(f, gps_por_caja.get(caja_n)))
        else:
            otros_carriers.append(_armar_fila(f, None))

    # Unidades del GPS que NO tienen fila abierta en el Excel -> "falta rellenar en Excel"
    gps_sin_excel: list[dict] = []
    for u in unidades_gps:
        caja_n = normalizar_clave(u.get("unidad"))
        if caja_n and caja_n not in cajas_excel_edr:
            gps_sin_excel.append({
                "caja": u.get("unidad", ""),
                "placas": u.get("placas", ""),
                "gps": {
                    "posicion": u.get("posicion"),
                    "velocidad": u.get("velocidad"),
                    "fechaHoraGps": u.get("fechaHoraGps"),
                    "direccion": u.get("direccion"),
                },
                "motivo": "Sin fila abierta en el Excel (falta rellenar o el embarque está cerrado)",
            })

    return {
        "edr": edr,
        "gpsSinExcel": gps_sin_excel,
        "otrosCarriers": otros_carriers,
        "totales": {
            "edr": len(edr),
            "gpsSinExcel": len(gps_sin_excel),
            "otrosCarriers": len(otros_carriers),
        },
    }


@app.patch("/api/embarques/fila/override")
async def patch_fila_override(
    body: FilaOverrideInput,
    fila_key_q: str,
    claims: dict = Depends(verificar_token),
):
    perfil = await get_perfil_usuario(claims, supabase_client)
    if perfil["cliente"] != "ADMIN":
        raise HTTPException(403, "Solo el administrador puede editar aquí")

    datos = {
        "fila_key": fila_key_q,
        "editado_por": claims.get("sub"),
        "editado_en": datetime.now(timezone.utc).isoformat(),
    }
    if body.trackingUrl is not None:
        datos["tracking_url"] = body.trackingUrl.strip()
    if body.comentario is not None:
        datos["comentario"] = body.comentario.strip()
    if body.loadOrden is not None:
        datos["load_orden"] = body.loadOrden.strip()

    supabase_client.table("embarque_fila_overrides").upsert(
        datos, on_conflict="fila_key"
    ).execute()
    await cargar_fila_overrides()
    return {"ok": True, "filaKey": fila_key_q}

@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket, token: str = ""):
    # 1. Validar token ANTES de aceptar la conexión
    try:
        claims = await verificar_token_ws(token)
        perfil = await get_perfil_usuario(claims, supabase_client)
        cliente = perfil["cliente"]
        rol = perfil["rol"]
    except Exception:
        await ws.close(code=1008)  # 1008 = policy violation
        return

    # 2. Conectar con su cliente
    await manager.connect(ws, cliente, rol)
    try:
        # snapshot inicial, ya filtrado
        todas = merge_con_meta(state.unidades_gps)
        await ws.send_json({
            "type": "snapshot",
            "timestamp": state.last_update,
            "unidades": filtrar_por_cliente(todas, cliente, rol),
        })
        while True:
            await ws.receive_text()
    except WebSocketDisconnect:
        await manager.disconnect(ws)
    except Exception:
        await manager.disconnect(ws)


@app.get("/api/health")
async def health():
    return {
        "status": "ok" if state.last_error is None else "degraded",
        "pollCount": state.poll_count,
        "lastUpdate": state.last_update,
        "lastError": state.last_error,
        "dashboardsConectados": len(manager.active),
    }



@app.post("/api/registro")
async def registro(body: RegistroInput):
    # 1. Validar el código de invitación (ahora trae cliente Y rol)
    res = supabase_client.table("codigos_invitacion").select("cliente,rol,activo").eq("codigo", body.codigo.strip()).execute()
    if not res.data or not res.data[0]["activo"]:
        raise HTTPException(400, "Código de invitación inválido")
    cliente = res.data[0]["cliente"]
    rol = res.data[0]["rol"]

    # 2. Crear el usuario en Supabase Auth (admin, con service_role)
    try:
        nuevo = supabase_client.auth.admin.create_user({
            "email": body.email.strip().lower(),
            "password": body.password,
            "email_confirm": True,
        })
        user_id = nuevo.user.id
    except Exception as e:
        raise HTTPException(400, f"No se pudo crear el usuario: {e}")

    # 3. Crear su perfil con el cliente Y rol del código
    supabase_client.table("perfiles").insert({"user_id": user_id, "cliente": cliente, "rol": rol}).execute()

    return {"ok": True, "cliente": cliente, "rol": rol}