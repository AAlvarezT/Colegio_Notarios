"""
Regression tests for the ODE accounting ETL pipeline.
Run from the repo root: pytest -q
"""
from pathlib import Path

import pandas as pd
import openpyxl

from transformer import (
    OUTPUT_COLUMNS,
    auto_create_bank_config,
    build_carga_dataframe,
    build_conciliacion,
    detect_total_row,
    load_ode_report,
    normalize_amount,
    normalize_document_value,
    normalize_serie,
    parse_date_value,
    process_report,
    read_raw_ode_report,
    serialize_serie_num,
    split_ode_rows,
)

ROOT = Path(__file__).resolve().parents[1]
SOURCE_PATH = ROOT / 'ETL_ODE' / 'reporte de cancelados agosto 2026.xlsx'
MODEL_PATH = ROOT / 'ETL_ODE' / 'FACTURA 07-2026 SHIKINA.xls'


# ---------------------------------------------------------------------------
# 1. Raw split — counts match the UI requirements exactly
# ---------------------------------------------------------------------------

def test_raw_split_counts():
    raw = read_raw_ode_report(SOURCE_PATH)
    raw, excluded, processed = split_ode_rows(raw)
    assert len(raw) == 151, f'expected 151 raw, got {len(raw)}'
    assert len(excluded) == 10, f'expected 10 excluded, got {len(excluded)}'
    assert len(processed) == 141, f'expected 141 processed, got {len(processed)}'


def test_raw_split_total_matches():
    raw = read_raw_ode_report(SOURCE_PATH)
    _, _, processed = split_ode_rows(raw)
    total = round(float(processed['importe_cancelado'].sum()), 2)
    assert total == 26757.00, f'expected S/26,757.00, got {total}'


# ---------------------------------------------------------------------------
# 2. Exclusion
# ---------------------------------------------------------------------------

def test_exclusion_serie_0001():
    df = load_ode_report(SOURCE_PATH)
    assert not df['serie'].str.startswith('0001').any()


# ---------------------------------------------------------------------------
# 3. D-row header fields must be blank
# ---------------------------------------------------------------------------

def test_d_row_document_fields_blank():
    df = load_ode_report(SOURCE_PATH)
    carga = build_carga_dataframe(df, periodo='202608', config={})
    d_rows = carga[carga['DEBE HABER'] == 'D']
    assert d_rows['TIP_ANEXO'].astype(str).str.strip().eq('').all(),  'TIP_ANEXO should be blank on D rows'
    assert d_rows['COD_ANEXO'].astype(str).str.strip().eq('').all(),  'COD_ANEXO should be blank on D rows'
    assert d_rows['TIP_DOC'].astype(str).str.strip().eq('').all(),    'TIP_DOC should be blank on D rows'
    assert d_rows['SERIE_NUM'].astype(str).str.strip().eq('').all(),  'SERIE_NUM should be blank on D rows'
    assert d_rows['FEC_VENC'].astype(str).str.strip().eq('').all(),   'FEC_VENC should be blank on D rows'
    assert d_rows['TIP_CAMBIO'].astype(str).str.strip().eq('').all(), 'TIP_CAMBIO should be blank on D rows'
    assert d_rows['CENTRO_COSTOS'].astype(str).str.strip().eq('').all(), 'CENTRO_COSTOS should be blank on D rows'
    assert d_rows['NRO_FILE'].astype(str).str.strip().eq('').all(),   'NRO_FILE should be blank on D rows'


def test_d_row_fec_doc_is_month_end():
    df = load_ode_report(SOURCE_PATH)
    carga = build_carga_dataframe(df, periodo='202608', config={})
    d_rows = carga[carga['DEBE HABER'] == 'D']
    for fec in d_rows['FEC_DOC']:
        dt = pd.to_datetime(fec, errors='coerce')
        assert dt.day == 31 and dt.month == 8 and dt.year == 2026, f'D-row FEC_DOC must be 31/08/2026, got {fec}'


# ---------------------------------------------------------------------------
# 4. Conciliation — exact expected values
# ---------------------------------------------------------------------------

EXPECTED_CONC = [
    ('INTERBANK',        '0001', 79,  20002.69, 20002.69, 0.00),
    ('BBVA CONTINENTAL', '0002', 15,  1720.00,  1720.00,  0.00),
    ('DE LA NACION',     '0003', 22,  4942.00,  4942.00,  0.00),
    ('OTROS',            '0004', 25,    92.31,    92.31,  0.00),
    ('TOTAL',           'TOTAL', 141, 26757.00, 26757.00, 0.00),
]


def test_conciliacion_exact_values():
    df = load_ode_report(SOURCE_PATH)
    carga = build_carga_dataframe(df, periodo='202608', config={})
    conc = build_conciliacion(carga)
    assert len(conc) == 5, f'expected 5 conciliation rows (4 banks + total), got {len(conc)}'
    for i, (banco, comp, cantidad, debito, credito, diferencia) in enumerate(EXPECTED_CONC):
        row = conc[i]
        assert row['banco'] == banco, f'row {i} banco: expected {banco!r}, got {row["banco"]!r}'
        assert row['comprobante'] == comp, f'row {i} comprobante: expected {comp!r}, got {row["comprobante"]!r}'
        assert row['cantidad'] == cantidad, f'row {i} cantidad: expected {cantidad}, got {row["cantidad"]}'
        assert abs(row['debito'] - debito) <= 0.01, f'row {i} debito: expected {debito}, got {row["debito"]}'
        assert abs(row['credito'] - credito) <= 0.01, f'row {i} credito: expected {credito}, got {row["credito"]}'
        assert abs(row['diferencia'] - diferencia) <= 0.01, f'row {i} diferencia: expected {diferencia}, got {row["diferencia"]}'
        assert row['estado'] == 'CUADRADO', f'row {i} estado should be CUADRADO'


# ---------------------------------------------------------------------------
# 5. Output workbook: exactly 22 columns, no _banco, formats applied
# ---------------------------------------------------------------------------

def test_output_has_exactly_22_columns(tmp_path):
    result = process_report(
        SOURCE_PATH, tmp_path, config={},
        fecha_calculo='fecha_cancelacion', periodo='202608',
        model_path=MODEL_PATH,
    )
    wb = openpyxl.load_workbook(result['output_file'])
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    assert headers == list(OUTPUT_COLUMNS), f'column mismatch: {headers}'
    assert '_banco' not in headers, '_banco aux column must not appear in exported file'


def test_output_date_and_number_formats(tmp_path):
    result = process_report(
        SOURCE_PATH, tmp_path, config={},
        fecha_calculo='fecha_cancelacion', periodo='202608',
    )
    wb = openpyxl.load_workbook(result['output_file'])
    ws = wb.active
    # Check first data row (row 2); D-row FEC_DOC is col E, IMPORT_TOTAL is col L
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.column_letter in ('E', 'J', 'N') and cell.value not in ('', None):
                assert 'D' in cell.number_format.upper() or 'Y' in cell.number_format.upper(), \
                    f'Expected date format on {cell.coordinate}, got {cell.number_format!r}'
            if cell.column_letter == 'L' and cell.value not in ('', None):
                assert '0' in cell.number_format, \
                    f'Expected numeric format on {cell.coordinate}, got {cell.number_format!r}'
        break  # one row is enough to verify formats were applied


# ---------------------------------------------------------------------------
# 6. Full process_report end-to-end
# ---------------------------------------------------------------------------

def test_process_report_real_report_generates_valid_carga(tmp_path):
    result = process_report(
        SOURCE_PATH, tmp_path, config={},
        fecha_calculo='fecha_cancelacion', periodo='202608',
        model_path=MODEL_PATH,
    )
    s = result['summary']
    assert s['cantidad_original'] == 151
    assert s['cantidad_excluida'] == 10
    assert s['cantidad_procesada'] == 141
    assert abs(s['total_procesado'] - 26757.00) <= 0.01
    assert abs(s['total_debito'] - 26757.00) <= 0.01
    assert abs(s['total_credito'] - 26757.00) <= 0.01
    assert s['diferencia'] == 0.00
    assert s['resultado'] == 'CUADRADO'

    carga = result['output_df']
    # 141 H rows + 4 D rows = 145
    assert len(carga) == 145, f'expected 145 output rows, got {len(carga)}'
    assert Path(result['output_file']).exists()
    assert Path(result['report_file']).exists()


# ---------------------------------------------------------------------------
# 7. analyze_source helper mirrors UI counts
# ---------------------------------------------------------------------------

def test_analyze_source_counts():
    from app import analyze_source
    raw, excluded, processed = analyze_source(str(SOURCE_PATH))
    assert len(raw) == 151
    assert len(excluded) == 10
    assert len(processed) == 141
    assert abs(float(processed['importe_cancelado'].sum()) - 26757.00) <= 0.01


# ---------------------------------------------------------------------------
# 8. Existing unit tests
# ---------------------------------------------------------------------------

def test_concatenacion_serie_num():
    assert serialize_serie_num({'serie': 'FE01', 'documento': '0029546'}) == 'FE010029546'


def test_totales_por_banco():
    df = load_ode_report(SOURCE_PATH)
    totals = df.groupby('banco')['importe_cancelado'].sum()
    for b in ('INTERBANK', 'BBVA CONTINENTAL', 'DE LA NACION', 'OTROS'):
        assert b in totals.index


def test_lectura_fechas_como_texto():
    dt = parse_date_value('25/08/2026')
    assert dt is not None and dt.day == 25 and dt.month == 8 and dt.year == 2026


def test_lectura_importes_con_comas():
    assert abs(normalize_amount('2,50') - 2.5) < 1e-9
    assert abs(normalize_amount('1.234,56') - 1234.56) < 1e-9
    assert abs(normalize_amount('1,234.56') - 1234.56) < 1e-9


def test_preservacion_ruc_y_ceros():
    assert normalize_document_value('02012345678') == '02012345678'
    assert normalize_document_value('000045') == '000045'
    assert normalize_serie('FE01') == 'FE01'


def test_deteccion_fila_de_totales():
    assert detect_total_row('TOTAL') is True
    assert detect_total_row('Reporte de Comprobantes Cancelados') is False


def test_creacion_automatica_banco_desconocido():
    cfg = {'bancos': {}}
    bank = auto_create_bank_config(cfg, 'BANCO NUEVO')
    assert bank['cuenta_contable'] == '10100011'
    assert bank['medio_pago'] == '008'


def test_exportacion_sin_nan_ni_nat():
    df = load_ode_report(SOURCE_PATH)
    out = build_carga_dataframe(df, periodo='202608', config={})
    assert not out[OUTPUT_COLUMNS].isnull().values.any()
    text = out[OUTPUT_COLUMNS].astype(str)
    assert not text.apply(lambda col: col.str.lower().str.contains('nan')).any().any()
    assert not text.apply(lambda col: col.str.lower().str.contains('nat')).any().any()


def test_model_file_uses_expected_structure():
    model = pd.read_excel(MODEL_PATH, header=0)
    assert list(model.columns) == list(OUTPUT_COLUMNS)


def test_igualdad_debe_haber():
    df = load_ode_report(SOURCE_PATH)
    out = build_carga_dataframe(df, periodo='202608', config={})
    debito = out[out['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()
    credito = out[out['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()
    assert abs(debito - credito) <= 0.01
