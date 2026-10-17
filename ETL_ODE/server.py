"""FastAPI web server for the Conversor ODE tool.

Entrypoint for Vercel: server:app
"""
from __future__ import annotations

import base64
import json
import logging
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
from fastapi import Cookie, FastAPI, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from transformer import (  # noqa: E402
    DEFAULT_CONFIG,
    build_carga_dataframe,
    build_conciliacion,
    create_validation_report,
    ensure_no_nan,
    export_carga_workbook,
    export_carga_xls,
    process_report,
    read_raw_ode_report,
    resolve_tip_doc_for_df,
    split_ode_rows,
    validate_report,
    validate_serie_map,
)

# ---------------------------------------------------------------------------
# Configuration from environment
# ---------------------------------------------------------------------------
SECRET_KEY: str = os.environ.get('SECRET_KEY', '')
ODE_USERNAME: str = os.environ.get('ODE_USERNAME', '')
ODE_PASSWORD: str = os.environ.get('ODE_PASSWORD', '')
IS_VERCEL: bool = bool(os.environ.get('VERCEL_ENV', ''))
# In local dev without credentials configured, allow access (shows warning)
DEV_MODE: bool = not IS_VERCEL and not (ODE_USERNAME and ODE_PASSWORD)

SESSION_MAX_AGE = 8 * 3600          # 8 hours in seconds
MAX_FILE_BYTES = 8 * 1024 * 1024    # 8 MB per uploaded file

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
app = FastAPI(title='Conversor ODE')

static_dir = HERE / 'static'
static_dir.mkdir(exist_ok=True)
app.mount('/static', StaticFiles(directory=str(static_dir)), name='static')

_log = logging.getLogger('conversor_ode')

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def _serializer() -> URLSafeTimedSerializer:
    key = SECRET_KEY or 'dev-insecure-key-change-in-prod'
    return URLSafeTimedSerializer(key)


def _sign_session(username: str) -> str:
    return _serializer().dumps({'u': username})


def _verify_session(token: str) -> Optional[str]:
    try:
        data = _serializer().loads(token, max_age=SESSION_MAX_AGE)
        return data.get('u')
    except (SignatureExpired, BadSignature):
        return None


def _get_session_user(request: Request) -> Optional[str]:
    token = request.cookies.get('ode_session')
    if not token:
        return None
    return _verify_session(token)


def _require_auth(request: Request) -> str:
    """Return username or raise HTTP 401. Dev mode bypasses auth."""
    if DEV_MODE:
        return 'dev'
    user = _get_session_user(request)
    if not user:
        raise HTTPException(status_code=401, detail='No autenticado')
    return user


def _check_prod_config() -> None:
    """In production (Vercel), raise 503 if required env vars are absent."""
    if IS_VERCEL and not (SECRET_KEY and ODE_USERNAME and ODE_PASSWORD):
        raise HTTPException(
            status_code=503,
            detail='Configuración incompleta en el servidor: faltan variables de entorno.'
        )

# ---------------------------------------------------------------------------
# JSON serialization helpers
# ---------------------------------------------------------------------------

def _safe_value(v: Any) -> Any:
    """Recursively convert non-JSON-safe values to safe equivalents."""
    if v is None:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime('%d/%m/%Y')
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
        return None
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        if np.isnan(v) or np.isinf(v):
            return None
        return float(v)
    if isinstance(v, np.bool_):
        return bool(v)
    if isinstance(v, dict):
        return {k: _safe_value(val) for k, val in v.items()}
    if isinstance(v, list):
        return [_safe_value(i) for i in v]
    if isinstance(v, str) and v.lower() in ('nan', 'nat', 'none'):
        return None
    return v


def _safe_json_response(data: Any) -> JSONResponse:
    return JSONResponse(content=_safe_value(data))

# ---------------------------------------------------------------------------
# Upload helpers
# ---------------------------------------------------------------------------

async def _read_upload(file: UploadFile, label: str) -> bytes:
    data = await file.read(MAX_FILE_BYTES + 1)
    if len(data) > MAX_FILE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f'{label}: el archivo supera el límite de 8 MB.'
        )
    return data


def _check_extension(filename: str, allowed: tuple) -> None:
    ext = Path(filename).suffix.lower()
    if ext not in allowed:
        raise HTTPException(
            status_code=422,
            detail=f'Tipo de archivo no permitido: {ext!r}. Se aceptan: {", ".join(allowed)}'
        )


def _build_config(config_json: str) -> Dict[str, Any]:
    try:
        user_cfg = json.loads(config_json) if config_json.strip() else {}
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=422, detail=f'config_json inválido: {exc}')
    merged: Dict[str, Any] = {**DEFAULT_CONFIG, **user_cfg}
    merged['bancos'] = {**DEFAULT_CONFIG['bancos'], **user_cfg.get('bancos', {})}
    return merged


def _get_periodo(year: str, month: str) -> str:
    try:
        return f'{int(year):04d}{int(month):02d}'
    except (ValueError, TypeError):
        raise HTTPException(status_code=422, detail='Año o mes inválido.')

# ---------------------------------------------------------------------------
# Analysis helpers
# ---------------------------------------------------------------------------

def _build_analyze_response(
    raw: pd.DataFrame,
    excluded: pd.DataFrame,
    processed: pd.DataFrame,
    periodo: str,
    config: Dict[str, Any],
    fecha_calculo: str,
) -> Dict[str, Any]:
    errors, warnings, info = validate_report(
        processed, periodo=periodo, fecha_calculo=fecha_calculo, config=config
    )

    tip_doc_series, tip_doc_obs = resolve_tip_doc_for_df(processed, config)

    # Preview (up to 200 rows)
    preview: List[Dict[str, Any]] = []
    fecha_col = {
        'fecha_cancelacion': 'fecha_cancelacion',
        'fecha_emi': 'fecha_emi',
    }.get(fecha_calculo)

    for (idx, row), tip in zip(processed.head(200).iterrows(), tip_doc_series.head(200)):
        fecha_val = ''
        if fecha_col and fecha_col in row.index:
            fv = row[fecha_col]
            if fv is not None and not (isinstance(fv, float) and np.isnan(fv)):
                try:
                    fecha_val = pd.to_datetime(fv).strftime('%d/%m/%Y')
                except Exception:
                    pass
        importe = row.get('importe_cancelado', 0.0)
        if importe is None or (isinstance(importe, float) and np.isnan(importe)):
            importe = 0.0
        preview.append({
            'banco': str(row.get('banco', '')),
            'serie': str(row.get('serie', '')),
            'documento': str(row.get('documento', '')),
            'importe': round(float(importe), 2),
            'fecha': fecha_val,
            'tip_doc': str(tip),
        })

    # Type summary
    ic = processed['importe_cancelado'] if 'importe_cancelado' in processed.columns else pd.Series(
        [0.0] * len(processed), index=processed.index
    )
    type_summary = {
        'facturas': int((tip_doc_series == '01').sum()),
        'boletas': int((tip_doc_series == '03').sum()),
        'sin_resolver': int((tip_doc_series == '').sum()),
        'importe_facturas': round(float(ic[tip_doc_series == '01'].sum()), 2),
        'importe_boletas': round(float(ic[tip_doc_series == '03'].sum()), 2),
        'importe_sin_resolver': round(float(ic[tip_doc_series == ''].sum()), 2),
    }

    # Series summary
    series_all: List[Dict[str, Any]] = []
    unresolved_series: List[str] = []
    if 'serie' in processed.columns:
        for serie in sorted(processed['serie'].astype(str).unique()):
            mask = processed['serie'].astype(str) == serie
            tipos_raw = list(tip_doc_series[mask].unique())
            tipos = [t for t in tipos_raw if t]
            resuelto = bool(tipos and all(t in ('01', '03') for t in tipos))
            series_all.append({'serie': serie, 'tipos': tipos, 'resuelto': resuelto})
            if not resuelto:
                unresolved_series.append(serie)

    # Observations (no server paths)
    obs_clean = [
        {
            'fila': obs.get('fila'),
            'serie': str(obs.get('serie', '')),
            'documento': str(obs.get('documento', '')),
            'tipo': str(obs.get('tipo', '')),
            'mensaje': str(obs.get('mensaje', '')),
        }
        for obs in tip_doc_obs
    ]

    stats = {
        'cantidad_original': len(raw),
        'cantidad_excluida': len(excluded),
        'cantidad_procesada': len(processed),
        'total_procesado': round(
            float(processed['importe_cancelado'].sum())
            if 'importe_cancelado' in processed.columns else 0.0,
            2,
        ),
    }

    return {
        'errors': errors,
        'warnings': warnings,
        'info': info,
        'preview': preview,
        'type_summary': type_summary,
        'series_all': series_all,
        'unresolved_series': unresolved_series,
        'tip_doc_obs': obs_clean,
        'stats': stats,
    }

# ---------------------------------------------------------------------------
# Routes — pages
# ---------------------------------------------------------------------------

@app.get('/', include_in_schema=False)
async def index_page(request: Request):
    _check_prod_config()
    user = _get_session_user(request)
    if not user and not DEV_MODE:
        return RedirectResponse(url='/login', status_code=302)
    html = static_dir / 'index.html'
    if not html.exists():
        return HTMLResponse('<h1>index.html no encontrado</h1>', status_code=500)
    return FileResponse(str(html), media_type='text/html')


@app.get('/login', include_in_schema=False)
async def login_page():
    html = static_dir / 'login.html'
    if not html.exists():
        return HTMLResponse(
            '<form method=post action=/login>'
            '<input name=username placeholder=Usuario>'
            '<input name=password type=password placeholder=Contraseña>'
            '<button>Ingresar</button></form>'
        )
    return FileResponse(str(html), media_type='text/html')

# ---------------------------------------------------------------------------
# Routes — auth
# ---------------------------------------------------------------------------

@app.post('/login')
async def login(
    username: str = Form(...),
    password: str = Form(...),
):
    _check_prod_config()
    if not DEV_MODE:
        if username != ODE_USERNAME or password != ODE_PASSWORD:
            raise HTTPException(status_code=401, detail='Credenciales incorrectas')
    token = _sign_session(username)
    resp = JSONResponse({'ok': True})
    resp.set_cookie(
        key='ode_session',
        value=token,
        httponly=True,
        samesite='lax',
        secure=IS_VERCEL,
        max_age=SESSION_MAX_AGE,
    )
    return resp


@app.post('/logout')
async def logout():
    resp = JSONResponse({'ok': True})
    resp.delete_cookie('ode_session')
    return resp

# ---------------------------------------------------------------------------
# Routes — config API
# ---------------------------------------------------------------------------

@app.get('/api/config/default')
async def config_default(request: Request):
    _check_prod_config()
    _require_auth(request)
    return _safe_json_response(DEFAULT_CONFIG)


@app.post('/api/config/validate')
async def config_validate(request: Request, config_json: str = Form(...)):
    _check_prod_config()
    _require_auth(request)
    try:
        cfg = json.loads(config_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=422, detail=f'JSON inválido: {exc}')

    serie_map = cfg.get('serie_map', {})
    if not isinstance(serie_map, dict):
        raise HTTPException(status_code=422, detail="'serie_map' debe ser un objeto JSON.")

    err = validate_serie_map(serie_map)
    if err:
        raise HTTPException(status_code=422, detail=err)

    required = ['cuenta_clientes', 'sub_diario', 'tipo_anexo', 'moneda', 'tipo_conversion', 'doc_anulado', 'glosa']
    for field in required:
        if field not in cfg:
            raise HTTPException(status_code=422, detail=f'Campo requerido faltante: {field}')

    return JSONResponse({'ok': True, 'message': 'Configuración válida'})

# ---------------------------------------------------------------------------
# Routes — processing API
# ---------------------------------------------------------------------------

@app.post('/api/analyze')
async def analyze(
    request: Request,
    reporte: UploadFile,
    year: str = Form(...),
    month: str = Form(...),
    fecha_calculo: str = Form('fecha_cancelacion'),
    config_json: str = Form('{}'),
    modelo: Optional[UploadFile] = None,
):
    _check_prod_config()
    _require_auth(request)

    _check_extension(reporte.filename or 'file.xlsx', ('.xlsx', '.xls'))
    reporte_bytes = await _read_upload(reporte, 'reporte')

    modelo_bytes: Optional[bytes] = None
    if modelo and modelo.filename and modelo.filename.strip():
        _check_extension(modelo.filename, ('.xlsx', '.xls'))
        modelo_bytes = await _read_upload(modelo, 'modelo')

    config = _build_config(config_json)
    periodo = _get_periodo(year, month)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            reporte_path = tmp / (Path(reporte.filename or 'reporte.xlsx').name)
            reporte_path.write_bytes(reporte_bytes)

            if modelo_bytes:
                modelo_path = tmp / (Path(modelo.filename or 'modelo.xlsx').name)
                modelo_path.write_bytes(modelo_bytes)

            raw = read_raw_ode_report(reporte_path)
            raw, excluded, processed = split_ode_rows(raw)
            result = _build_analyze_response(raw, excluded, processed, periodo, config, fecha_calculo)
    except HTTPException:
        raise
    except Exception as exc:
        _log.error('analyze error: %s', exc, exc_info=True)
        raise HTTPException(status_code=422, detail=str(exc))

    return _safe_json_response(result)


@app.post('/api/generate')
async def generate(
    request: Request,
    reporte: UploadFile,
    year: str = Form(...),
    month: str = Form(...),
    fecha_calculo: str = Form('fecha_cancelacion'),
    config_json: str = Form('{}'),
    modelo: Optional[UploadFile] = None,
):
    _check_prod_config()
    _require_auth(request)

    _check_extension(reporte.filename or 'file.xlsx', ('.xlsx', '.xls'))
    reporte_bytes = await _read_upload(reporte, 'reporte')

    modelo_bytes: Optional[bytes] = None
    if modelo and modelo.filename and modelo.filename.strip():
        _check_extension(modelo.filename, ('.xlsx', '.xls'))
        modelo_bytes = await _read_upload(modelo, 'modelo')

    config = _build_config(config_json)
    periodo = _get_periodo(year, month)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            reporte_path = tmp / (Path(reporte.filename or 'reporte.xlsx').name)
            reporte_path.write_bytes(reporte_bytes)

            modelo_path = None
            if modelo_bytes:
                modelo_path = tmp / (Path(modelo.filename or 'modelo.xlsx').name)
                modelo_path.write_bytes(modelo_bytes)

            out_dir = tmp / 'salida'
            out_dir.mkdir()

            result = process_report(
                source_path=reporte_path,
                output_dir=out_dir,
                config=config,
                fecha_calculo=fecha_calculo,
                periodo=periodo,
                model_path=modelo_path,
            )

            xls_bytes = Path(result['output_file']).read_bytes()
            xlsx_bytes = Path(result['output_file_xlsx']).read_bytes()
            rpt_bytes = Path(result['report_file']).read_bytes()

            xls_name = Path(result['output_file']).name
            xlsx_name = Path(result['output_file_xlsx']).name
            rpt_name = Path(result['report_file']).name

        files = [
            {
                'name': xls_name,
                'mime': 'application/vnd.ms-excel',
                'data': base64.b64encode(xls_bytes).decode('ascii'),
                'label': 'Archivo para macro (.xls)',
            },
            {
                'name': xlsx_name,
                'mime': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'data': base64.b64encode(xlsx_bytes).decode('ascii'),
                'label': 'Copia de respaldo (.xlsx)',
            },
            {
                'name': rpt_name,
                'mime': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'data': base64.b64encode(rpt_bytes).decode('ascii'),
                'label': 'Reporte de validación (solo revisión)',
            },
        ]

        summary = _safe_value(result['summary'])
        summary.pop('archivo_origen', None)  # don't expose server paths

        return _safe_json_response({'files': files, 'summary': summary})

    except HTTPException:
        raise
    except Exception as exc:
        _log.error('generate error: %s', exc, exc_info=True)
        raise HTTPException(status_code=422, detail=str(exc))
