from __future__ import annotations

from datetime import datetime

import openpyxl
import pandas as pd
import pytest

from transformer import (
    DEFAULT_CONFIG,
    build_carga_dataframe,
    process_report,
    validate_report,
)


def _cfg() -> dict:
    cfg = {**DEFAULT_CONFIG}
    cfg['serie_map'] = {'FE01': '01'}
    return cfg


def _df() -> pd.DataFrame:
    return pd.DataFrame([
        {
            'numruc': '20111111111',
            'serie': 'FE01',
            'documento': '000001',
            'importe_cancelado': 100.0,
            'banco': 'INTERBANK',
            'fecha_ingreso': datetime(2026, 8, 5),
            'fecha_cancelacion': datetime(2026, 8, 5),
            'fecha_emi': datetime(2026, 7, 1),
            'tipo_comprobante': '01',
            '_source_row': 4,
        },
        {
            'numruc': '20111111111',
            'serie': 'FE01',
            'documento': '000001',
            'importe_cancelado': 40.0,
            'banco': 'INTERBANK',
            'fecha_ingreso': datetime(2026, 8, 8),
            'fecha_cancelacion': datetime(2026, 8, 8),
            'fecha_emi': datetime(2026, 7, 1),
            'tipo_comprobante': '01',
            '_source_row': 5,
        },
        {
            'numruc': '20222222222',
            'serie': 'FE01',
            'documento': '000002',
            'importe_cancelado': 60.0,
            'banco': 'INTERBANK',
            'fecha_ingreso': datetime(2026, 8, 8),
            'fecha_cancelacion': datetime(2026, 8, 8),
            'fecha_emi': datetime(2026, 8, 2),
            'tipo_comprobante': '01',
            '_source_row': 6,
        },
    ])


def test_grouping_splits_by_collection_date_and_groups_same_day():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg())
    d_rows = carga[carga['DEBE HABER'] == 'D']
    h_rows = carga[carga['DEBE HABER'] == 'H']
    assert len(d_rows) == 2
    assert len(h_rows) == 3

    by_comp = h_rows.groupby('COMPROBANTE').size().to_dict()
    assert sorted(by_comp.values()) == [1, 2]


def test_partial_payments_are_kept_as_independent_h_rows():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg())
    h_rows = carga[carga['DEBE HABER'] == 'H']
    same_doc = h_rows[
        h_rows['FEC_DOC'].apply(lambda v: pd.to_datetime(v).date() == datetime(2026, 7, 1).date())
        & h_rows['IMPORT_TOTAL'].isin([100.0, 40.0])
    ]
    assert len(same_doc) == 2


def test_document_date_kept_and_fec_reg_uses_collection_date():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg())
    h_rows = carga[carga['DEBE HABER'] == 'H'].copy()
    row = h_rows[h_rows['IMPORT_TOTAL'] == 100.0].iloc[0]
    assert pd.to_datetime(row['FEC_DOC']).date() == datetime(2026, 7, 1).date()
    assert pd.to_datetime(row['FEC_REG']).date() in {
        datetime(2026, 8, 5).date(),
        datetime(2026, 8, 8).date(),
    }


def test_bank_and_detail_lines_have_coherent_dates():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg())
    for comp, grp in carga.groupby('COMPROBANTE'):
        d_row = grp[grp['DEBE HABER'] == 'D'].iloc[0]
        f_reg = pd.to_datetime(d_row['FEC_REG']).date()
        assert pd.to_datetime(d_row['FEC_DOC']).date() == f_reg
        assert grp['FEC_REG'].apply(lambda v: pd.to_datetime(v).date()).eq(f_reg).all(), comp


def test_balance_per_asiento_and_file():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg())
    for _, grp in carga.groupby('COMPROBANTE'):
        debe = float(grp[grp['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum())
        haber = float(grp[grp['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum())
        assert round(debe - haber, 2) == 0.0
    total_debe = float(carga[carga['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum())
    total_haber = float(carga[carga['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum())
    assert round(total_debe - total_haber, 2) == 0.0


def test_invalid_or_missing_collection_date_blocks():
    bad = _df().copy()
    bad.loc[1, 'fecha_ingreso'] = None
    errors, _, _ = validate_report(bad, periodo='202608', config=_cfg())
    assert any('fecha de cobro' in err.lower() for err in errors)


def test_range_filter_is_inclusive():
    carga = build_carga_dataframe(
        _df(),
        periodo='202608',
        config=_cfg(),
        fecha_desde='2026-08-05',
        fecha_hasta='2026-08-08',
    )
    assert int((carga['DEBE HABER'] == 'H').sum()) == 3


def test_cross_month_selection_is_blocked():
    mixed = _df().copy()
    mixed.loc[2, 'fecha_ingreso'] = datetime(2026, 9, 1)
    with pytest.raises(ValueError, match='más de un mes'):
        build_carga_dataframe(mixed, periodo='202608', config=_cfg())


def test_correlativo_inicial_is_respected():
    carga = build_carga_dataframe(_df(), periodo='202608', config=_cfg(), correlativo_inicial=7)
    comps = sorted({int(c) for c in carga['COMPROBANTE'].astype(str).unique()})
    assert comps[0] == 7


def test_fixed_series_rule_blocks_wrong_map():
    errors, _, _ = validate_report(
        _df(),
        periodo='202608',
        config={**_cfg(), 'serie_map': {'FE01': '03'}},
    )
    assert any("Serie 'FE01'" in e and 'esperado 01' in e for e in errors)


def test_fixed_series_rule_blocks_wrong_explicit_column():
    bad = _df().copy()
    bad.loc[0, 'tipo_comprobante'] = '03'
    errors, _, _ = validate_report(
        bad,
        periodo='202608',
        config={**_cfg(), 'serie_map': {}},
    )
    assert any('tipo_comprobante=03' in e and "Serie 'FE01'" in e for e in errors)


def test_process_report_uses_only_in_range_movements(tmp_path):
    src = tmp_path / 'rango.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['NumRuc', 'Serie', 'Documento', 'Importe Cancelado', 'Fecha Ingreso', 'Fecha Cancelacion', 'Banco', 'Tipo Comprobante'])
    ws.append(['20111111111', 'FE01', '000001', 100.0, datetime(2026, 8, 5), datetime(2026, 8, 5), 'INTERBANK', '01'])
    ws.append(['20111111112', 'FE01', '000002', 100.0, datetime(2026, 8, 20), datetime(2026, 8, 20), 'INTERBANK', '01'])
    wb.save(src)

    result = process_report(
        src,
        tmp_path,
        config={**_cfg(), 'serie_map': {}},
        periodo='202608',
        fecha_desde='2026-08-01',
        fecha_hasta='2026-08-10',
    )
    summary = result['summary']
    assert summary['total_reporte'] == 200.0
    assert summary['total_procesado'] == 100.0
    assert summary['cantidad_reporte_total'] == 2
    assert summary['cantidad_procesada'] == 1
    assert summary['cantidad_fuera_rango'] == 1
    assert int((result['output_df']['DEBE HABER'] == 'H').sum()) == 1
