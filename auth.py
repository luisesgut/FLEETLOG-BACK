"""
auth.py — Verificación de JWT de Supabase (asimétrico ES256 vía JWKS)
y resolución del cliente del usuario.
"""
import os
import time
import httpx
from dotenv import load_dotenv   # <-- agregar
from fastapi import Header, HTTPException
from jose import jwt

load_dotenv()   # <-- agregar, ANTES de leer las variables

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")  # service_role, ya lo tienes
JWKS_URL = f"{SUPABASE_URL}/auth/v1/.well-known/jwks.json"

# Cache del JWKS (se refresca cada 10 min; así no pegamos a Supabase en cada request)
_jwks_cache: dict = {"keys": None, "ts": 0}
_JWKS_TTL = 600  # 10 minutos


async def _get_jwks() -> dict:
    now = time.time()
    if _jwks_cache["keys"] is None or (now - _jwks_cache["ts"]) > _JWKS_TTL:
        async with httpx.AsyncClient(timeout=10) as client:
            r = await client.get(JWKS_URL)
            r.raise_for_status()
            _jwks_cache["keys"] = r.json()
            _jwks_cache["ts"] = now
    return _jwks_cache["keys"]


async def verificar_token(authorization: str = Header(default="")) -> dict:
    """Valida el JWT de Supabase y regresa sus claims. Lanza 401 si es inválido."""
    if not authorization.startswith("Bearer "):
        raise HTTPException(401, "Falta token de autorización")
    token = authorization.split(" ", 1)[1]

    try:
        jwks = await _get_jwks()
        claims = jwt.decode(
            token,
            jwks,                       # python-jose acepta el JWKS dict directo
            algorithms=["ES256"],
            audience="authenticated",   # Supabase pone este audience
        )
        return claims
    except Exception as e:
        raise HTTPException(401, f"Token inválido: {e}")


async def get_cliente_usuario(claims: dict, supabase_client) -> str:
    """Del token saca el user_id, busca su cliente en 'perfiles'.
    Regresa el cliente (ej 'DESTINY') o 'ADMIN'."""
    user_id = claims.get("sub")
    if not user_id:
        raise HTTPException(401, "Token sin user_id")

    res = supabase_client.table("perfiles").select("cliente").eq("user_id", user_id).execute()
    if not res.data:
        raise HTTPException(403, "Usuario sin cliente asignado")
    return res.data[0]["cliente"]

async def verificar_token_ws(token: str) -> dict:
    """Igual que verificar_token pero recibe el token crudo (del query param del WS)."""
    if not token:
        raise HTTPException(401, "Falta token")
    try:
        jwks = await _get_jwks()
        return jwt.decode(token, jwks, algorithms=["ES256"], audience="authenticated")
    except Exception as e:
        raise HTTPException(401, f"Token inválido: {e}")