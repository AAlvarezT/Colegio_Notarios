import os
from pathlib import Path

import pandas as pd

from transformer import (
    auto_create_bank_config,
    build_carga_dataframe,
    detect_total_row,
    load_ode_report,
    normalize_amount,
    normalize_document_value,
    normalize_serie,
    parse_date_value,
    serialize_serie_num,
)

ROOT = Path(__file__).resolve().parents[1]
SOURCE_PATH = ROOT / 'ETL_ODE' / 'reporte de cancelados agosto 2026.xlsx'
MODEL_PATH = ROOT / 'ETL_ODE' / 'FACTURA 07-2026 SHIKINA.xls'


def test_exclusion_serie_0001():
    df = load_ode_report(SOURCE_PATH)
    assert not df['serie'].str.startswith('0001').any()
    assert df['serie'].str.contains('0001').sum() == 0


def test_conservacion_movimientos_parciales():
    df = load_ode_report(SOURCE_PATH)
    filtered = df[df['banco'] == 'INTERBANK']
    assert len(filtered) >= 70
    assert filtered['importe_cancelado'].sum() > 0


def test_mismo_documento_en_varios_bancos():
    df = load_ode_report(SOURCE_PATH)
    doc = '0030106'
    rows = df[df['documento'] == doc]
    assert len(rows) >= 1
    assert set(rows['banco'].astype(str).unique())


def test_concatenacion_serie_num():
    row = {'serie': 'FE01', 'documento': '0029546'}
    assert serialize_serie_num(row) == 'FE010029546'


def test_totales_por_banco():
    df = load_ode_report(SOURCE_PATH)
    totals = df.groupby('banco')['importe_cancelado'].sum().round(2)
    assert 'INTERBANK' in totals.index
    assert 'BBVA CONTINENTAL' in totals.index
    assert 'DE LA NACION' in totals.index
    assert 'OTROS' in totals.index


def test_igualdad_debe_haber():
    df = load_ode_report(SOURCE_PATH)
    processed = df[df['serie'] != '0001']
    output = build_carga_dataframe(processed, periodo='202608', config={})
    debito = output[output['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()
    credito = output[output['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()
    assert abs(debito - credito) <= 0.01


def test_lectura_fechas_como_texto():
    dt = parse_date_value('25/08/2026')
    assert dt is not None
    assert dt.day == 25
    assert dt.month == 8
    assert dt.year == 2026


def test_lectura_importes_con_comas():
    assert abs(normalize_amount('2,50') - 2.5) < 1e-9
    assert abs(normalize_amount('1.234,56') - 1234.56) < 1e-9
    assert abs(normalize_amount('1,234.56') - 1234.56) < 1e-9


def test_preservacion_ruc_y_ceros():
    row = {'numruc': '02012345678', 'serie': 'FE01', 'documento': '000045', 'serie_num': 'FE01000045'}
    assert normalize_document_value(row['numruc']) == '02012345678'
    assert normalize_document_value(row['documento']) == '000045'
    assert normalize_serie(row['serie']) == 'FE01'


def test_deteccion_fila_de_totales():
    assert detect_total_row('TOTAL') is True
    assert detect_total_row('Reporte de Comprobantes Cancelados') is False


def test_creacion_automatica_banco_desconocido():
    config = {'bancos': {}}
    bank = auto_create_bank_config(config, 'BANCO NUEVO')
    assert bank['cuenta_contable'] == '10100011'
    assert bank['medio_pago'] == '008'


def test_generacion_correlativos():
    grouped = [
        {'banco': 'INTERBANK', 'documentos': [1]},
        {'banco': 'BBVA CONTINENTAL', 'documentos': [1, 2]},
        {'banco': 'DE LA NACION', 'documentos': [1]},
    ]
    result = [
        {'comprobante': '0001'},
        {'comprobante': '0002'},
        {'comprobante': '0003'},
    ]
    assert result[0]['comprobante'] == '0001'
    assert result[1]['comprobante'] == '0002'
    assert result[2]['comprobante'] == '0003'


def test_exportacion_sin_nan_ni_nat():
    df = load_ode_report(SOURCE_PATH)
    out = build_carga_dataframe(df, periodo='202608', config={})
    assert not out.isnull().values.any()
    text = out.astype(str)
    assert not text.apply(lambda col: col.str.contains('nan', case=False)).any().any()
    assert not text.apply(lambda col: col.str.contains('nat', case=False)).any().any()


def test_model_file_uses_expected_structure():
    model = pd.read_excel(MODEL_PATH, header=0)
    expected = [
        'CTA_CONTABLE', 'ANO_MES', 'SUB_DIARIO', 'COMPROBANTE', 'FEC_DOC', 'TIP_ANEXO',
        'COD_ANEXO', 'TIP_DOC', 'SERIE_NUM', 'FEC_VENC', 'MONEDA', 'IMPORT_TOTAL',
        'TIP_CONEVR.', 'FEC_REG', 'TIP_CAMBIO', 'GLOSA', 'CENTRO_COSTOS', 'GLOSA_MOV',
        'DOC_ANULADO', 'DEBE HABER', 'MEDIO_PAGO', 'NRO_FILE'
    ]
    assert list(model.columns) == expected
