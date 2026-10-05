from __future__ import annotations

import json
import logging
import re
import unicodedata
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

OUTPUT_COLUMNS = [
    'CTA_CONTABLE', 'ANO_MES', 'SUB_DIARIO', 'COMPROBANTE', 'FEC_DOC', 'TIP_ANEXO',
    'COD_ANEXO', 'TIP_DOC', 'SERIE_NUM', 'FEC_VENC', 'MONEDA', 'IMPORT_TOTAL',
    'TIP_CONEVR.', 'FEC_REG', 'TIP_CAMBIO', 'GLOSA', 'CENTRO_COSTOS', 'GLOSA_MOV',
    'DOC_ANULADO', 'DEBE HABER', 'MEDIO_PAGO', 'NRO_FILE'
]

BANK_PRIORITY = ['INTERBANK', 'BBVA CONTINENTAL', 'DE LA NACION', 'OTROS']
DEFAULT_CONFIG = {
    'cuenta_clientes': '12120001',
    'sub_diario': '01',
    'tipo_anexo': '02',
    'tipo_documento': '01',
    'moneda': 'MN',
    'tipo_conversion': 'VTA',
    'doc_anulado': '0',
    'glosa': 'COBRANZA DEL MES',
    'fecha_documento': 'fecha_cancelacion',
    'bancos': {
        'INTERBANK': {'cuenta_contable': '10100011', 'medio_pago': '008'},
        'BBVA CONTINENTAL': {'cuenta_contable': '10100011', 'medio_pago': '008'},
        'DE LA NACION': {'cuenta_contable': '10100011', 'medio_pago': '008'},
        'OTROS': {'cuenta_contable': '10100011', 'medio_pago': '008'},
    }
}

LOGGER = logging.getLogger('conversor_ode')
if not LOGGER.handlers:
    log_dir = Path(__file__).resolve().parent / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)
    LOGGER.setLevel(logging.INFO)
    handler = logging.FileHandler(log_dir / 'conversor_ode.log', encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
    handler.setFormatter(formatter)
    LOGGER.addHandler(handler)


def normalize_header_name(value: Any) -> str:
    cleaned = normalize_text(value)
    cleaned = unicodedata.normalize('NFKD', cleaned).encode('ascii', 'ignore').decode('utf-8')
    cleaned = cleaned.lower().replace(' ', '_').replace('-', '_').replace('.', '').replace('/', '_')
    cleaned = re.sub(r'_+', '_', cleaned).strip('_')
    return cleaned


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime('%d/%m/%Y')
    if isinstance(value, (float, np.floating)):
        if np.isnan(value):
            return ''
        if value.is_integer():
            return str(int(value))
        return str(value)
    text = str(value).strip()
    return text


def detect_total_row(value: Any) -> bool:
    text = normalize_text(value).upper()
    return text.startswith('TOTAL') or 'TOTAL' in text


def _remove_accents(value: str) -> str:
    return unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('utf-8')


def normalize_bank_name(value: Any) -> str:
    text = normalize_text(value).upper().strip()
    if not text:
        return ''
    text = _remove_accents(text)
    if 'INTERBANK' in text:
        return 'INTERBANK'
    if 'BBVA' in text or 'CONTINENTAL' in text:
        return 'BBVA CONTINENTAL'
    if 'NACION' in text or 'NACIÓN' in text:
        return 'DE LA NACION'
    if 'OTROS' in text:
        return 'OTROS'
    return text.title()


def normalize_serie(value: Any) -> str:
    text = normalize_text(value).upper().strip()
    return _remove_accents(text)


def normalize_document_value(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    text = normalize_text(value)
    if not text:
        return ''
    if isinstance(value, (float, np.floating)):
        return format(float(value), 'f').rstrip('0').rstrip('.')
    if 'E-' in text.upper() or 'E+' in text.upper():
        try:
            return format(float(text), 'f').rstrip('0').rstrip('.')
        except Exception:
            pass
    return text


def serialize_serie_num(row: Dict[str, Any]) -> str:
    serie = normalize_serie(row.get('serie', ''))
    documento = normalize_document_value(row.get('documento', ''))
    if not serie and not documento:
        return ''
    if not documento:
        return serie
    if not serie:
        return documento
    return f'{serie}{documento}'


def normalize_amount(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = normalize_text(value).replace('\u00a0', '').replace(' ', '')
    if not text or text.lower() in {'nan', 'nat', 'none', ''}:
        return 0.0
    if text.startswith('(') and text.endswith(')'):
        text = f'-{text[1:-1]}'
    text = text.replace('S/', '').replace('$', '')
    if '.' in text and ',' in text:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '').replace(',', '.')
        else:
            text = text.replace(',', '')
    elif ',' in text:
        if text.count(',') > 1:
            text = text.replace(',', '')
        else:
            left, right = text.split(',', 1)
            if len(right) in (2, 3):
                if len(right) == 3 and left.isdigit() and int(left) > 999:
                    text = f'{left}{right}'
                else:
                    text = f'{left}.{right}'
            else:
                text = f'{left}.{right}'
    elif '.' in text:
        if text.count('.') > 1:
            parts = text.split('.')
            if len(parts[-1]) == 3:
                text = ''.join(parts)
    try:
        return float(text)
    except Exception:
        return 0.0


def parse_date_value(value: Any) -> Optional[datetime]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            return pd.to_datetime(value, unit='D', origin='1899-12-30', errors='coerce').to_pydatetime()
        except Exception:
            return None
    text = normalize_text(value)
    if not text:
        return None
    text = text.replace(' 00:00:00', '')
    text = text.replace('T', ' ')
    for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y', '%m/%d/%Y', '%m/%d/%y']:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(text, errors='coerce').to_pydatetime()
    except Exception:
        return None


def detect_header_row(raw_df: pd.DataFrame) -> int:
    for idx in range(min(raw_df.shape[0], 20)):
        row = raw_df.iloc[idx].fillna('')
        text_values = [normalize_header_name(v) for v in row.tolist()]
        score = 0
        for token in ['numruc', 'serie', 'documento', 'importe_cancelado', 'fecha_cancelacion', 'banco']:
            if token in text_values:
                score += 1
        if score >= 3:
            return idx
    raise ValueError('No se pudo localizar la fila de encabezado del reporte de ODE.')


def read_raw_ode_report(path: str | Path) -> pd.DataFrame:
    source_path = Path(path)
    if not source_path.exists():
        raise FileNotFoundError(f'No se encontró el archivo: {source_path}')

    try:
        df = pd.read_excel(source_path, header=None)
    except Exception:
        try:
            df = pd.read_excel(source_path, header=None, engine='xlrd')
        except Exception as exc:
            raise ValueError(f'No se pudo leer el archivo de origen: {exc}') from exc

    header_row = detect_header_row(df)
    headers = []
    for cell in df.iloc[header_row].tolist():
        headers.append(normalize_header_name(cell))
    data = df.iloc[header_row + 1:].copy()
    data.columns = headers
    data = data.dropna(how='all').reset_index(drop=True)

    if not data.empty:
        last_row = data.iloc[-1]
        last_text = ' '.join([normalize_text(v) for v in last_row.fillna('').tolist()])
        if 'TOTAL' in last_text.upper():
            data = data.iloc[:-1].copy()

    cleaned = data.copy()
    for c in cleaned.columns:
        cleaned[c] = cleaned[c].fillna('')

    required = {
        'tipo_movi': 'tipo_movi',
        'fecha_ingreso': 'fecha_ingreso',
        'grupo_cli': 'grupo_cli',
        'nombre_cli': 'nombre_cli',
        'numruc': 'numruc',
        'fecha_emi': 'fecha_emi',
        'serie': 'serie',
        'documento': 'documento',
        'kardex': 'kardex',
        'fecha_cancelacion': 'fecha_cancelacion',
        'banco': 'banco',
        'nro_cuenta_banco': 'nro_cuenta_banco',
        'movimiento': 'movimiento',
        'importe_notarial': 'importe_notarial',
        'importe_registral': 'importe_registral',
        'importe_cancelado': 'importe_cancelado',
        'saldo': 'saldo',
        'fecha_ven': 'fecha_ven',
        'monedatipo': 'monedatipo',
        'tipopago': 'tipopago',
        'codigoemp': 'codigoemp',
        'comentarios': 'comentarios',
    }
    rename_map = {k: v for k, v in required.items() if k in cleaned.columns}
    cleaned = cleaned.rename(columns=rename_map)

    if 'numruc' in cleaned.columns:
        cleaned['numruc'] = cleaned['numruc'].apply(normalize_document_value)
    if 'documento' in cleaned.columns:
        cleaned['documento'] = cleaned['documento'].apply(normalize_document_value)
    if 'serie' in cleaned.columns:
        cleaned['serie'] = cleaned['serie'].apply(normalize_serie)
    if 'banco' in cleaned.columns:
        cleaned['banco'] = cleaned['banco'].apply(normalize_bank_name)
    if 'importe_cancelado' in cleaned.columns:
        cleaned['importe_cancelado'] = cleaned['importe_cancelado'].apply(normalize_amount)
    if 'importe_notarial' in cleaned.columns:
        cleaned['importe_notarial'] = cleaned['importe_notarial'].apply(normalize_amount)
    if 'importe_registral' in cleaned.columns:
        cleaned['importe_registral'] = cleaned['importe_registral'].apply(normalize_amount)
    if 'saldo' in cleaned.columns:
        cleaned['saldo'] = cleaned['saldo'].apply(normalize_amount)
    if 'fecha_cancelacion' in cleaned.columns:
        cleaned['fecha_cancelacion'] = cleaned['fecha_cancelacion'].apply(parse_date_value)
    if 'fecha_emi' in cleaned.columns:
        cleaned['fecha_emi'] = cleaned['fecha_emi'].apply(parse_date_value)
    if 'fecha_ven' in cleaned.columns:
        cleaned['fecha_ven'] = cleaned['fecha_ven'].apply(parse_date_value)
    if 'fecha_ingreso' in cleaned.columns:
        cleaned['fecha_ingreso'] = cleaned['fecha_ingreso'].apply(parse_date_value)

    cleaned = cleaned[cleaned.apply(lambda row: not row.astype(str).str.contains(r'^\s*$', regex=True).all(), axis=1)]
    cleaned = cleaned.reset_index(drop=True)
    return cleaned


def load_ode_report(path: str | Path) -> pd.DataFrame:
    df = read_raw_ode_report(path)
    if 'serie' in df.columns:
        df = df[~df['serie'].astype(str).str.startswith('0001')].copy()
    return df.reset_index(drop=True)


def validate_report(df: pd.DataFrame) -> Tuple[List[str], List[str], List[str]]:
    errors: List[str] = []
    warnings: List[str] = []
    info: List[str] = []

    if df.empty:
        errors.append('El archivo no contiene movimientos válidos.')
        return errors, warnings, info

    required = ['numruc', 'serie', 'documento', 'importe_cancelado', 'banco', 'fecha_cancelacion']
    for col in required:
        if col not in df.columns:
            errors.append(f'Falta la columna requerida: {col}')

    if 'serie' in df.columns:
        bad_serie = df[df['serie'].astype(str).str.strip() == '']
        if not bad_serie.empty:
            errors.append('Hay filas con serie vacía.')

    if 'documento' in df.columns:
        bad_doc = df[df['documento'].astype(str).str.strip() == '']
        if not bad_doc.empty:
            errors.append('Hay filas con documento vacío.')

    if 'numruc' in df.columns:
        bad_ruc = df[df['numruc'].astype(str).str.len().lt(8)]
        if not bad_ruc.empty:
            errors.append('Hay filas con RUC demasiado corto o vacío.')

    if 'importe_cancelado' in df.columns:
        bad_amount = df[df['importe_cancelado'] <= 0]
        if not bad_amount.empty:
            errors.append('Hay movimientos con importe_cancelado vacío o menor o igual a cero.')

    if 'fecha_cancelacion' in df.columns:
        bad_dates = df[df['fecha_cancelacion'].isna()]
        if not bad_dates.empty:
            errors.append('Hay movimientos con fecha de cancelación inválida.')

    if 'banco' in df.columns:
        bank_empty = df[df['banco'].astype(str).str.strip() == '']
        if not bank_empty.empty:
            errors.append('Hay movimientos sin nombre de banco.')

    if 'serie' in df.columns:
        excluded = df[df['serie'].astype(str).str.startswith('0001')]
        if not excluded.empty:
            info.append(f'Se excluyeron {len(excluded)} movimientos de serie 0001.')

    if 'banco' in df.columns and 'OTROS' in df['banco'].unique():
        warnings.append('Los movimientos OTROS pueden corresponder a una retención del 3%. Verifique la cuenta contable configurada.')

    if not errors:
        info.append('Archivo legible y con estructura válida.')

    return errors, warnings, info


def auto_create_bank_config(config: Dict[str, Any], nombre_banco: str) -> Dict[str, Any]:
    config = config or {'bancos': {}}
    config.setdefault('bancos', {})
    bank_name = normalize_bank_name(nombre_banco)
    if bank_name not in config['bancos']:
        config['bancos'][bank_name] = {'cuenta_contable': '10100011', 'medio_pago': '008'}
    return config['bancos'][bank_name]


def load_config(path: str | Path) -> Dict[str, Any]:
    config_path = Path(path)
    if not config_path.exists():
        return DEFAULT_CONFIG.copy()
    with open(config_path, 'r', encoding='utf-8') as fh:
        loaded = json.load(fh)
    merged = DEFAULT_CONFIG.copy()
    merged.update(loaded)
    merged['bancos'] = {**DEFAULT_CONFIG['bancos'], **loaded.get('bancos', {})}
    return merged


def save_config(path: str | Path, config: Dict[str, Any]) -> None:
    config_path = Path(path)
    config_path.parent.mkdir(parents=True, exist_ok=True)
    with open(config_path, 'w', encoding='utf-8') as fh:
        json.dump(config, fh, ensure_ascii=False, indent=2)


def get_period_from_report(df: pd.DataFrame) -> str:
    candidate = ''
    if 'fecha_cancelacion' in df.columns and df['fecha_cancelacion'].notna().any():
        dates = df['fecha_cancelacion'].dropna()
        if not dates.empty:
            candidate = dates.iloc[0]
            if isinstance(candidate, pd.Timestamp):
                candidate = candidate.to_pydatetime()
            return f'{candidate.year:04d}{candidate.month:02d}'
    return datetime.now().strftime('%Y%m')


def build_carga_dataframe(processed_df: pd.DataFrame, periodo: str, config: Dict[str, Any]) -> pd.DataFrame:
    config = {**DEFAULT_CONFIG, **(config or {})}
    config['bancos'] = {**DEFAULT_CONFIG['bancos'], **config.get('bancos', {})}

    if processed_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    base_rows: List[Dict[str, Any]] = []
    ordered_banks = []
    seen = set()
    for bank in BANK_PRIORITY:
        if bank in processed_df['banco'].astype(str).unique():
            ordered_banks.append(bank)
            seen.add(bank)
    extra = sorted({b for b in processed_df['banco'].astype(str).unique() if b and b not in seen}, key=str.upper)
    ordered_banks.extend(extra)

    for bank in ordered_banks:
        bank_rows = processed_df[processed_df['banco'].astype(str) == bank]
        if bank_rows.empty:
            continue
        bank_cfg = config['bancos'].get(bank, {'cuenta_contable': '10100011', 'medio_pago': '008'})
        bank_account = bank_cfg.get('cuenta_contable', '10100011')
        medio = bank_cfg.get('medio_pago', '008')
        total = bank_rows['importe_cancelado'].sum()
        base_rows.append({
            'CTA_CONTABLE': str(bank_account),
            'ANO_MES': str(periodo),
            'SUB_DIARIO': str(config.get('sub_diario', '01')),
            'COMPROBANTE': str(len(base_rows) + 1).zfill(4),
            'FEC_DOC': pd.Timestamp(f'{periodo[:4]}-{periodo[4:6]}-{int(periodo[6:8]) if len(periodo) > 6 else 1}').to_pydatetime() if len(periodo) >= 6 else None,
            'TIP_ANEXO': '',
            'COD_ANEXO': '',
            'TIP_DOC': '',
            'SERIE_NUM': '',
            'FEC_VENC': '',
            'MONEDA': str(config.get('moneda', 'MN')),
            'IMPORT_TOTAL': round(float(total), 2),
            'TIP_CONEVR.': str(config.get('tipo_conversion', 'VTA')),
            'FEC_REG': pd.Timestamp(f'{periodo[:4]}-{periodo[4:6]}-{int(periodo[6:8]) if len(periodo) > 6 else 1}').to_pydatetime() if len(periodo) >= 6 else None,
            'TIP_CAMBIO': '',
            'GLOSA': f'COBRANZA DEL MES - {bank}',
            'CENTRO_COSTOS': '',
            'GLOSA_MOV': f'COBRANZA DEL MES - {bank}',
            'DOC_ANULADO': str(config.get('doc_anulado', '0')),
            'DEBE HABER': 'D',
            'MEDIO_PAGO': str(medio),
            'NRO_FILE': '',
            '_bank_key': bank,
            '_group_rows': bank_rows,
            '_comprobante': str(len(base_rows) + 1).zfill(4),
        })

    rows: List[Dict[str, Any]] = []
    for item in base_rows:
        bank = item['_bank_key']
        bank_rows = item['_group_rows']
        comprobante = item['_comprobante']
        for _, row in bank_rows.iterrows():
            fecha_doc = row.get('fecha_cancelacion') if not pd.isna(row.get('fecha_cancelacion')) else row.get('fecha_emi')
            fecha_venc = row.get('fecha_cancelacion') if not pd.isna(row.get('fecha_cancelacion')) else row.get('fecha_emi')
            rows.append({
                'CTA_CONTABLE': str(config.get('cuenta_clientes', '12120001')),
                'ANO_MES': str(periodo),
                'SUB_DIARIO': str(config.get('sub_diario', '01')),
                'COMPROBANTE': comprobante,
                'FEC_DOC': pd.Timestamp(fecha_doc).to_pydatetime() if pd.notna(fecha_doc) else '',
                'TIP_ANEXO': str(config.get('tipo_anexo', '02')),
                'COD_ANEXO': normalize_document_value(row.get('numruc', '')),
                'TIP_DOC': str(config.get('tipo_documento', '01')),
                'SERIE_NUM': serialize_serie_num({'serie': row.get('serie', ''), 'documento': row.get('documento', '')}),
                'FEC_VENC': pd.Timestamp(fecha_venc).to_pydatetime() if pd.notna(fecha_venc) else '',
                'MONEDA': str(config.get('moneda', 'MN')),
                'IMPORT_TOTAL': round(float(normalize_amount(row.get('importe_cancelado', 0))), 2),
                'TIP_CONEVR.': str(config.get('tipo_conversion', 'VTA')),
                'FEC_REG': pd.Timestamp(f'{periodo[:4]}-{periodo[4:6]}-{int(periodo[6:8]) if len(periodo) > 6 else 1}').to_pydatetime() if len(periodo) >= 6 else '',
                'TIP_CAMBIO': '',
                'GLOSA': str(config.get('glosa', 'COBRANZA DEL MES')),
                'CENTRO_COSTOS': '',
                'GLOSA_MOV': str(config.get('glosa', 'COBRANZA DEL MES')),
                'DOC_ANULADO': str(config.get('doc_anulado', '0')),
                'DEBE HABER': 'H',
                'MEDIO_PAGO': str(config['bancos'].get(bank, {'medio_pago': '008'}).get('medio_pago', '008')),
                'NRO_FILE': '',
            })
        rows.append({
            'CTA_CONTABLE': str(item['CTA_CONTABLE']),
            'ANO_MES': str(periodo),
            'SUB_DIARIO': str(item['SUB_DIARIO']),
            'COMPROBANTE': comprobante,
            'FEC_DOC': item['FEC_DOC'],
            'TIP_ANEXO': '',
            'COD_ANEXO': '',
            'TIP_DOC': '',
            'SERIE_NUM': '',
            'FEC_VENC': item['FEC_VENC'],
            'MONEDA': str(item['MONEDA']),
            'IMPORT_TOTAL': round(float(item['IMPORT_TOTAL']), 2),
            'TIP_CONEVR.': str(item['TIP_CONEVR.']),
            'FEC_REG': item['FEC_REG'],
            'TIP_CAMBIO': '',
            'GLOSA': str(item['GLOSA']),
            'CENTRO_COSTOS': '',
            'GLOSA_MOV': str(item['GLOSA_MOV']),
            'DOC_ANULADO': str(item['DOC_ANULADO']),
            'DEBE HABER': 'D',
            'MEDIO_PAGO': str(item['MEDIO_PAGO']),
            'NRO_FILE': '',
        })

    output = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    for column in ['COMPROBANTE', 'COD_ANEXO', 'SERIE_NUM', 'MONEDA', 'DOC_ANULADO', 'MEDIO_PAGO', 'ANO_MES']:
        output[column] = output[column].fillna('').astype(str)
    output = output.replace({np.nan: ''})
    output = output.fillna('')
    output['IMPORT_TOTAL'] = pd.to_numeric(output['IMPORT_TOTAL'], errors='coerce').fillna(0.0)
    return output


def create_validation_report(summary: Dict[str, Any], source_path: Path, output_path: Path) -> None:
    from openpyxl.styles import Font, PatternFill

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        summary_df = pd.DataFrame([
            ['Fecha de generación', summary['fecha_generacion']],
            ['Archivo de origen', str(source_path)],
            ['Periodo', summary['periodo']],
            ['Cantidad original', summary['cantidad_original']],
            ['Cantidad excluida', summary['cantidad_excluida']],
            ['Cantidad procesada', summary['cantidad_procesada']],
            ['Total procesado', summary['total_procesado']],
            ['Total débito', summary['total_debito']],
            ['Total crédito', summary['total_credito']],
            ['Diferencia', summary['diferencia']],
            ['Resultado', summary['resultado']],
        ], columns=['Campo', 'Valor'])
        summary_df.to_excel(writer, index=False, sheet_name='Resumen')

        processed_df = pd.DataFrame(summary['procesados'])
        if not processed_df.empty:
            processed_df.to_excel(writer, index=False, sheet_name='Procesados')
        else:
            pd.DataFrame([{'mensaje': 'Sin registros procesados'}]).to_excel(writer, index=False, sheet_name='Procesados')

        excluidos_df = pd.DataFrame(summary['excluidos'])
        if not excluidos_df.empty:
            excluidos_df.to_excel(writer, index=False, sheet_name='Excluidos')
        else:
            pd.DataFrame([{'mensaje': 'Sin exclusiones'}]).to_excel(writer, index=False, sheet_name='Excluidos')

        observ_df = pd.DataFrame(summary['observaciones'])
        if not observ_df.empty:
            observ_df.to_excel(writer, index=False, sheet_name='Observaciones')
        else:
            pd.DataFrame([{'mensaje': 'Sin observaciones'}]).to_excel(writer, index=False, sheet_name='Observaciones')

        conciliacion_df = pd.DataFrame(summary['conciliacion'])
        if not conciliacion_df.empty:
            conciliacion_df.to_excel(writer, index=False, sheet_name='Conciliacion')
        else:
            pd.DataFrame([{'banco': 'N/A', 'comprobante': 'N/A', 'cantidad': 0, 'debito': 0, 'credito': 0, 'diferencia': 0}]).to_excel(writer, index=False, sheet_name='Conciliacion')

        fill = PatternFill('solid', fgColor='1F2A44')
        font = Font(color='FFFFFF', bold=True)
        for ws in writer.sheets.values():
            ws.freeze_panes = 'A2'
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font


def export_carga_workbook(output_df: pd.DataFrame, output_path: Path, file_label: str = 'CARGA') -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = None
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        ws = wb.active
        ws.title = 'Hoja1'
        ws.freeze_panes = 'A2'
        ws.append(list(OUTPUT_COLUMNS))
        for _, row in output_df.iterrows():
            values = []
            for col in OUTPUT_COLUMNS:
                value = row.get(col, '')
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    values.append('')
                else:
                    values.append(value)
            ws.append(values)
        header_fill = PatternFill(fill_type='solid', fgColor='1F2A44')
        header_font = Font(color='FFFFFF', bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        for col in ws.columns:
            max_len = 0
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 22)
        ws.auto_filter.ref = ws.dimensions
        for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.column in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]:
                    pass
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=22):
            for cell in row:
                try:
                    if isinstance(cell.value, (datetime, date)):
                        cell.number_format = 'dd/mm/yyyy'
                    elif isinstance(cell.value, (int, float)):
                        if cell.column_letter in {'L', 'M', 'N'}:
                            cell.number_format = '0.00'
                except Exception:
                    pass
        wb.save(output_path)
    except Exception:
        if wb is not None:
            try:
                wb.save(output_path)
            except Exception:
                pass
        raise


def ensure_no_nan(output_df: pd.DataFrame) -> pd.DataFrame:
    for column in output_df.columns:
        output_df[column] = output_df[column].apply(lambda value: '' if value is None or (isinstance(value, float) and pd.isna(value)) or str(value).lower() in {'nan', 'nat'} else value)
    return output_df


def build_summary(processed_df: pd.DataFrame, excluded_df: pd.DataFrame, warnings: List[str], errors: List[str], info: List[str], source_path: Path, periodo: str) -> Dict[str, Any]:
    total_procesado = float(processed_df['importe_cancelado'].sum()) if 'importe_cancelado' in processed_df.columns else 0.0
    total_debito = float(processed_df['importe_cancelado'].sum()) if 'importe_cancelado' in processed_df.columns else 0.0
    total_credito = total_debito
    total_diff = abs(total_debito - total_credito)
    resultado = 'CUADRADO' if total_diff <= 0.01 else 'REVISAR'
    return {
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        'archivo_origen': str(source_path),
        'periodo': periodo,
        'cantidad_original': int(len(processed_df) + len(excluded_df)),
        'cantidad_excluida': int(len(excluded_df)),
        'cantidad_procesada': int(len(processed_df)),
        'total_procesado': round(total_procesado, 2),
        'total_debito': round(total_debito, 2),
        'total_credito': round(total_credito, 2),
        'diferencia': round(total_diff, 2),
        'resultado': resultado,
        'warnings': warnings,
        'errors': errors,
        'info': info,
        'procesados': processed_df.to_dict(orient='records'),
        'excluidos': excluded_df.to_dict(orient='records'),
        'conciliacion': [],
        'observaciones': [{'tipo': 'warning', 'mensaje': m} for m in warnings] + [{'tipo': 'info', 'mensaje': m} for m in info],
    }


def process_report(source_path: str | Path, output_dir: str | Path, config: Optional[Dict[str, Any]] = None, fecha_calculo: str = 'fecha_cancelacion') -> Dict[str, Any]:
    if config is None:
        config = load_config(Path(__file__).resolve().parent / 'config.json')
    source_path = Path(source_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    raw = read_raw_ode_report(source_path)
    if raw.empty:
        raise ValueError('No se encontraron movimientos válidos en el archivo de ODE.')

    excluded = raw[raw['serie'].astype(str).str.startswith('0001')].copy()
    processed = raw[~raw['serie'].astype(str).str.startswith('0001')].copy()
    for idx in processed.index:
        if pd.isna(processed.at[idx, 'banco']) or str(processed.at[idx, 'banco']).strip() == '':
            processed.at[idx, 'banco'] = 'OTROS'
    for idx in excluded.index:
        excluded.at[idx, 'motivo'] = 'Serie 0001 - operación registral excluida'

    if config:
        for bank_name in processed['banco'].dropna().unique():
            normalized = normalize_bank_name(bank_name)
            auto_create_bank_config(config, normalized)
    period = get_period_from_report(processed) or datetime.now().strftime('%Y%m')
    carga = build_carga_dataframe(processed, period, config)
    carga = ensure_no_nan(carga)

    if carga.empty:
        raise ValueError('No se pudieron generar registros contables a partir del reporte.')

    total_carga = float(carga['IMPORT_TOTAL'].sum()) if 'IMPORT_TOTAL' in carga.columns else 0.0
    debito = float(carga[carga['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()) if 'DEBE HABER' in carga.columns else 0.0
    credito = float(carga[carga['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()) if 'DEBE HABER' in carga.columns else 0.0
    diff = abs(debito - credito)
    resumen = {
        'archivo_origen': str(source_path),
        'periodo': period,
        'cantidad_original': len(raw),
        'cantidad_excluida': len(excluded),
        'cantidad_procesada': len(processed),
        'total_procesado': round(total_carga, 2),
        'total_debito': round(debito, 2),
        'total_credito': round(credito, 2),
        'diferencia': round(diff, 2),
        'resultado': 'CUADRADO' if diff <= 0.01 else 'REVISAR',
        'procesados': processed.to_dict(orient='records'),
        'excluidos': excluded.to_dict(orient='records'),
        'conciliacion': [],
    }

    out_file = output_dir / f'CARGA_COBRANZAS_ODE_{period}.xlsx'
    report_file = output_dir / f'REPORTE_VALIDACION_ODE_{period}.xlsx'
    if out_file.exists():
        out_file = output_dir / f'CARGA_COBRANZAS_ODE_{period}_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    if report_file.exists():
        report_file = output_dir / f'REPORTE_VALIDACION_ODE_{period}_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'

    export_carga_workbook(carga, out_file, file_label='CARGA')
    create_validation_report({
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        'periodo': period,
        'cantidad_original': len(raw),
        'cantidad_excluida': len(excluded),
        'cantidad_procesada': len(processed),
        'total_procesado': round(total_carga, 2),
        'total_debito': round(debito, 2),
        'total_credito': round(credito, 2),
        'diferencia': round(diff, 2),
        'resultado': 'CUADRADO' if diff <= 0.01 else 'REVISAR',
        'procesados': processed.to_dict(orient='records'),
        'excluidos': excluded.to_dict(orient='records'),
        'conciliacion': [],
        'observaciones': [{'tipo': 'warning', 'mensaje': 'Los movimientos OTROS pueden corresponder a una retención del 3%. Verifique la cuenta contable configurada.'}] if 'OTROS' in processed['banco'].astype(str).unique() else [],
    }, source_path, report_file)

    result = {
        'summary': resumen,
        'output_file': str(out_file),
        'report_file': str(report_file),
        'processed_df': processed,
        'excluded_df': excluded,
        'output_df': carga,
        'errors': [],
        'warnings': ['Los movimientos OTROS pueden corresponder a una retención del 3%. Verifique la cuenta contable configurada.'] if 'OTROS' in processed['banco'].astype(str).unique() else [],
        'info': [f'{len(processed)} movimientos procesados', f'{len(excluded)} excluidos'],
        'period': period,
    }
    return result
