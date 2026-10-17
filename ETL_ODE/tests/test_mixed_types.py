"""
Tests for mixed-type (Factura 01 + Boleta 03) support.

All scenarios use synthetic DataFrames; no real file is required.
"""
from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
import xlrd

from transformer import (
    DEFAULT_CONFIG,
    OUTPUT_COLUMNS,
    build_carga_dataframe,
    build_conciliacion,
    export_carga_xls,
    normalize_tip_doc_code,
    process_report,
    resolve_tip_doc,
    resolve_tip_doc_for_df,
    split_ode_rows,
    validate_report,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DATE = datetime(2026, 8, 15)


def _make_df(**overrides) -> pd.DataFrame:
    """Minimal valid processed DataFrame with two rows (FE01 + BE01)."""
    base = {
        'numruc': ['20000000001', '20000000002'],
        'serie': ['FE01', 'BE01'],
        'documento': ['0000001', '0000002'],
        'importe_cancelado': [100.0, 50.0],
        'banco': ['INTERBANK', 'INTERBANK'],
        'fecha_cancelacion': [_DATE, _DATE],
    }
    base.update(overrides)
    return pd.DataFrame(base)


def _cfg_serie_map(**kw) -> dict:
    return {**DEFAULT_CONFIG, 'serie_map': {'FE': '01', 'BE': '03', **kw}}


# ---------------------------------------------------------------------------
# 1. normalize_tip_doc_code
# ---------------------------------------------------------------------------

class TestNormalizeTipDocCode:
    def test_string_01(self):
        assert normalize_tip_doc_code('01') == '01'

    def test_string_03(self):
        assert normalize_tip_doc_code('03') == '03'

    def test_int_1(self):
        assert normalize_tip_doc_code(1) == '01'

    def test_int_3(self):
        assert normalize_tip_doc_code(3) == '03'

    def test_float_1(self):
        assert normalize_tip_doc_code(1.0) == '01'

    def test_float_3(self):
        assert normalize_tip_doc_code(3.0) == '03'

    def test_bare_string_1(self):
        assert normalize_tip_doc_code('1') == '01'

    def test_bare_string_3(self):
        assert normalize_tip_doc_code('3') == '03'

    def test_unknown_returns_none(self):
        assert normalize_tip_doc_code('02') is None
        assert normalize_tip_doc_code('FAC') is None
        assert normalize_tip_doc_code('') is None

    def test_none_returns_none(self):
        assert normalize_tip_doc_code(None) is None

    def test_nan_returns_none(self):
        assert normalize_tip_doc_code(float('nan')) is None


# ---------------------------------------------------------------------------
# 2. resolve_tip_doc – per-row logic
# ---------------------------------------------------------------------------

class TestResolveTipDoc:
    # 2a. Global fallback when serie_map is empty
    def test_global_fallback_01(self):
        cfg = {**DEFAULT_CONFIG, 'tipo_documento': '01', 'serie_map': {}}
        td, msg = resolve_tip_doc({'serie': 'F001'}, cfg)
        assert td == '01' and msg is None

    def test_global_fallback_03(self):
        cfg = {**DEFAULT_CONFIG, 'tipo_documento': '03', 'serie_map': {}}
        td, msg = resolve_tip_doc({'serie': 'B001'}, cfg)
        assert td == '03' and msg is None

    # 2b. Serie map prefix matching
    def test_serie_map_prefix_factura(self):
        td, msg = resolve_tip_doc({'serie': 'FE01'}, _cfg_serie_map())
        assert td == '01' and msg is None

    def test_serie_map_prefix_boleta(self):
        td, msg = resolve_tip_doc({'serie': 'BE01'}, _cfg_serie_map())
        assert td == '03' and msg is None

    def test_serie_map_exact_match(self):
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE01': '01'}}
        td, msg = resolve_tip_doc({'serie': 'FE01'}, cfg)
        assert td == '01' and msg is None

    # 2c. Explicit tipo_comprobante column
    def test_explicit_column_01(self):
        td, msg = resolve_tip_doc({'serie': 'X999', 'tipo_comprobante': '01'}, DEFAULT_CONFIG)
        assert td == '01' and msg is None

    def test_explicit_column_03(self):
        td, msg = resolve_tip_doc({'serie': 'X999', 'tipo_comprobante': '03'}, DEFAULT_CONFIG)
        assert td == '03' and msg is None

    def test_explicit_column_integer(self):
        td, msg = resolve_tip_doc({'serie': 'X999', 'tipo_comprobante': 3}, DEFAULT_CONFIG)
        assert td == '03' and msg is None

    # 2d. Unknown type → None + message
    def test_unknown_explicit_blocks(self):
        td, msg = resolve_tip_doc({'serie': 'X999', 'tipo_comprobante': '99'}, DEFAULT_CONFIG)
        assert td is None
        assert '99' in msg

    # 2e. Unknown serie with active map → None + message
    def test_unknown_serie_with_map_blocks(self):
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        td, msg = resolve_tip_doc({'serie': 'BE01'}, cfg)
        assert td is None
        assert 'BE01' in msg or 'BE' in msg

    # 2f. Conflict: explicit vs map → None + message
    def test_conflict_explicit_vs_map_blocks(self):
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        td, msg = resolve_tip_doc({'serie': 'FE01', 'tipo_comprobante': '03'}, cfg)
        assert td is None
        assert 'Conflicto' in msg or 'conflicto' in msg

    # 2g. serie_map with invalid code
    def test_invalid_code_in_map_blocks(self):
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '99'}}
        td, msg = resolve_tip_doc({'serie': 'FE01'}, cfg)
        assert td is None
        assert '99' in msg


# ---------------------------------------------------------------------------
# 3. resolve_tip_doc_for_df
# ---------------------------------------------------------------------------

class TestResolveTipDocForDf:
    def test_all_resolved_via_map(self):
        df = _make_df()
        series, obs = resolve_tip_doc_for_df(df, _cfg_serie_map())
        assert list(series) == ['01', '03']
        assert obs == []

    def test_partial_unresolved(self):
        df = _make_df(serie=['FE01', 'ZZ99'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        series, obs = resolve_tip_doc_for_df(df, cfg)
        assert series.iloc[0] == '01'
        assert series.iloc[1] == ''
        assert len(obs) == 1
        assert obs[0]['serie'] == 'ZZ99'

    def test_conflict_detected(self):
        df = _make_df(serie=['FE01', 'FE01'], tipo_comprobante=['01', '03'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        series, obs = resolve_tip_doc_for_df(df, cfg)
        # Row 1 ok (01 == map 01), row 2 conflict (03 vs map 01)
        assert series.iloc[0] == '01'
        assert series.iloc[1] == ''
        assert any(o['tipo'] == 'CONFLICTO' or 'Conflicto' in o['mensaje'] for o in obs)


# ---------------------------------------------------------------------------
# 4. validate_report – TIP_DOC blocking
# ---------------------------------------------------------------------------

class TestValidateReportTipDoc:
    def test_passes_with_valid_map(self):
        df = _make_df()
        errors, _, _ = validate_report(df, periodo='202608', config=_cfg_serie_map())
        assert errors == []

    def test_blocks_on_unresolved_serie(self):
        df = _make_df(serie=['FE01', 'ZZ99'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        errors, _, _ = validate_report(df, periodo='202608', config=cfg)
        assert any('ZZ99' in e for e in errors)

    def test_blocks_on_conflict(self):
        df = _make_df(serie=['FE01', 'FE01'], tipo_comprobante=['01', '03'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        errors, _, _ = validate_report(df, periodo='202608', config=cfg)
        assert any('Conflicto' in e or 'conflicto' in e for e in errors)

    def test_no_config_skips_tipDoc_check(self):
        """When config is not supplied, validate_report must NOT add TIP_DOC errors."""
        df = _make_df(serie=['ZZ99', 'ZZ88'])
        errors, _, _ = validate_report(df, periodo='202608')  # no config
        assert not any('tipo_comprobante' in e.lower() or 'tip_doc' in e.lower() for e in errors)


# ---------------------------------------------------------------------------
# 5. build_carga_dataframe – mixed types
# ---------------------------------------------------------------------------

class TestBuildCargaMixed:
    def _build(self, **overrides):
        df = _make_df(**overrides)
        return build_carga_dataframe(df, periodo='202608', config=_cfg_serie_map())

    def test_h_rows_have_correct_tip_doc(self):
        carga = self._build()
        h = carga[carga['DEBE HABER'] == 'H']
        assert list(h['TIP_DOC']) == ['01', '03']

    def test_d_rows_tip_doc_blank(self):
        carga = self._build()
        d = carga[carga['DEBE HABER'] == 'D']
        assert d['TIP_DOC'].astype(str).str.strip().eq('').all()

    def test_accounting_balance_mixed(self):
        carga = self._build()
        debito  = carga[carga['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()
        credito = carga[carga['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()
        assert abs(debito - credito) <= 0.01

    def test_raises_on_unresolved(self):
        df = _make_df(serie=['FE01', 'ZZ99'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {'FE': '01'}}
        with pytest.raises(ValueError, match='ZZ99'):
            build_carga_dataframe(df, periodo='202608', config=cfg)

    def test_global_fallback_single_type(self):
        """When serie_map is empty, tipo_documento='01' applies to all rows."""
        df = _make_df(serie=['FE01', 'BE01'])
        cfg = {**DEFAULT_CONFIG, 'tipo_documento': '01', 'serie_map': {}}
        carga = build_carga_dataframe(df, periodo='202608', config=cfg)
        h = carga[carga['DEBE HABER'] == 'H']
        assert all(v == '01' for v in h['TIP_DOC'])

    def test_explicit_column_overrides(self):
        df = _make_df(tipo_comprobante=['01', '03'])
        cfg = {**DEFAULT_CONFIG, 'serie_map': {}}
        carga = build_carga_dataframe(df, periodo='202608', config=cfg)
        h = carga[carga['DEBE HABER'] == 'H']
        assert list(h['TIP_DOC']) == ['01', '03']


# ---------------------------------------------------------------------------
# 6. 0001 exclusion preserved
# ---------------------------------------------------------------------------

def test_exclusion_0001_preserved():
    df = pd.DataFrame({
        'numruc': ['20000000001', '20000000002', '20000000003'],
        'serie':  ['0001', 'FE01', 'BE01'],
        'documento': ['0001', '0000002', '0000003'],
        'importe_cancelado': [200.0, 100.0, 50.0],
        'banco': ['INTERBANK', 'INTERBANK', 'INTERBANK'],
        'fecha_cancelacion': [_DATE, _DATE, _DATE],
    })
    _, excluded, processed = split_ode_rows(df)
    assert len(excluded) == 1
    assert len(processed) == 2
    assert not processed['serie'].str.zfill(4).eq('0001').any()


# ---------------------------------------------------------------------------
# 7. Partial payments / duplicates preserved
# ---------------------------------------------------------------------------

def test_partial_payments_not_merged():
    """Two rows with same serie+documento but different amounts must both be kept."""
    df = _make_df(
        serie=['FE01', 'FE01'],
        documento=['0000001', '0000001'],
        importe_cancelado=[80.0, 20.0],
    )
    carga = build_carga_dataframe(df, periodo='202608', config=_cfg_serie_map())
    h = carga[carga['DEBE HABER'] == 'H']
    assert len(h) == 2
    assert abs(h['IMPORT_TOTAL'].sum() - 100.0) <= 0.01


# ---------------------------------------------------------------------------
# 8. XLS export – TIP_DOC and leading zeros
# ---------------------------------------------------------------------------

def test_xls_tip_doc_values_and_text_format(tmp_path):
    df = _make_df()
    carga = build_carga_dataframe(df, periodo='202608', config=_cfg_serie_map())
    xls_path = tmp_path / 'test_mixed.xls'
    export_carga_xls(carga, xls_path)

    wb = xlrd.open_workbook(str(xls_path))
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    tip_col = headers.index('TIP_DOC')
    dh_col  = headers.index('DEBE HABER')

    h_types = []
    for r in range(1, ws.nrows):
        if str(ws.cell_value(r, dh_col)) == 'H':
            cell = ws.cell(r, tip_col)
            # Must be TEXT type (xlrd ctype 1) to preserve leading zero
            assert cell.ctype == xlrd.XL_CELL_TEXT, (
                f'TIP_DOC at row {r+1} must be TEXT, got ctype={cell.ctype}'
            )
            h_types.append(cell.value)

    assert '01' in h_types
    assert '03' in h_types


def test_xls_tip_doc_leading_zero_preserved(tmp_path):
    """'01' and '03' must be stored as the strings '01' and '03', not 1 or 3."""
    df = _make_df()
    carga = build_carga_dataframe(df, periodo='202608', config=_cfg_serie_map())
    xls_path = tmp_path / 'test_leading_zero.xls'
    export_carga_xls(carga, xls_path)

    wb = xlrd.open_workbook(str(xls_path))
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    tip_col = headers.index('TIP_DOC')
    dh_col  = headers.index('DEBE HABER')

    for r in range(1, ws.nrows):
        if str(ws.cell_value(r, dh_col)) == 'H':
            val = ws.cell_value(r, tip_col)
            assert val in ('01', '03'), (
                f'TIP_DOC at row {r+1} should be "01" or "03", got {val!r}'
            )


# ---------------------------------------------------------------------------
# 9. process_report end-to-end with synthetic file
# ---------------------------------------------------------------------------

def _build_synthetic_excel(path: Path) -> None:
    """Write a minimal ODE-style Excel that passes all existing validators."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([
        'NumRuc', 'Serie', 'Documento', 'Importe Cancelado',
        'Fecha Cancelacion', 'Banco', 'Tipo Comprobante',
    ])
    rows = [
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK', '01'),
        ('20222222222', 'BE01', '0000002',  50.0, datetime(2026, 8, 11), 'INTERBANK', '03'),
        ('20333333333', 'FE01', '0000003',  75.0, datetime(2026, 8, 12), 'OTROS',     '01'),
    ]
    for r in rows:
        ws.append(r)
    wb.save(str(path))


def test_process_report_mixed_synthetic(tmp_path):
    source = tmp_path / 'synth.xlsx'
    _build_synthetic_excel(source)
    result = process_report(
        source, tmp_path,
        config={**DEFAULT_CONFIG, 'serie_map': {}},
        fecha_calculo='fecha_cancelacion',
        periodo='202608',
    )
    s = result['summary']
    assert s['cantidad_procesada'] == 3
    assert abs(s['total_procesado'] - 225.0) <= 0.01
    assert s['resultado'] == 'CUADRADO'

    ts = s['tip_doc_summary']
    assert ts['01'] == 2
    assert ts['03'] == 1
    assert ts['sin_resolver'] == 0

    # Verify in exported .xls
    wb = xlrd.open_workbook(result['output_file'])
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    tip_col = headers.index('TIP_DOC')
    dh_col  = headers.index('DEBE HABER')
    tip_vals = [ws.cell_value(r, tip_col) for r in range(1, ws.nrows)
                if str(ws.cell_value(r, dh_col)) == 'H']
    assert '01' in tip_vals
    assert '03' in tip_vals


# ---------------------------------------------------------------------------
# 10. Accounting balance for mixed report
# ---------------------------------------------------------------------------

def test_accounting_balance_mixed_multibank():
    df = pd.DataFrame({
        'numruc': ['20000000001', '20000000002', '20000000003'],
        'serie':  ['FE01', 'BE01', 'FE01'],
        'documento': ['001', '002', '003'],
        'importe_cancelado': [100.0, 50.0, 75.0],
        'banco': ['INTERBANK', 'BBVA CONTINENTAL', 'OTROS'],
        'fecha_cancelacion': [_DATE, _DATE, _DATE],
    })
    carga = build_carga_dataframe(df, periodo='202608', config=_cfg_serie_map())
    debito  = carga[carga['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()
    credito = carga[carga['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()
    assert abs(debito - credito) <= 0.01
    assert abs(debito - 225.0) <= 0.01

    conc = build_conciliacion(carga)
    total_row = next(r for r in conc if r['banco'] == 'TOTAL')
    assert total_row['estado'] == 'CUADRADO'


# ---------------------------------------------------------------------------
# 11. Validation report contains TipoDoc sheet
# ---------------------------------------------------------------------------

def test_validation_report_has_tipodoc_sheet(tmp_path):
    import openpyxl
    source = tmp_path / 'synth2.xlsx'
    _build_synthetic_excel(source)
    result = process_report(source, tmp_path, config={**DEFAULT_CONFIG, 'serie_map': {}})
    wb = openpyxl.load_workbook(result['report_file'])
    assert 'TipoDoc' in wb.sheetnames
    ws = wb['TipoDoc']
    counts = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}
    assert counts.get('Facturas (01)') == 2
    assert counts.get('Boletas (03)') == 1
    assert counts.get('Sin resolver') == 0


# ---------------------------------------------------------------------------
# 12. Single-type report still works (backward compatibility)
# ---------------------------------------------------------------------------

def test_single_type_report_backward_compat(tmp_path):
    """An all-facturas report with no serie_map and tipo_documento='01' must still generate."""
    df = _make_df(serie=['FE01', 'FE01'], documento=['0001', '0002'])
    cfg = {**DEFAULT_CONFIG, 'tipo_documento': '01', 'serie_map': {}}
    carga = build_carga_dataframe(df, periodo='202608', config=cfg)
    h = carga[carga['DEBE HABER'] == 'H']
    assert all(v == '01' for v in h['TIP_DOC'])
    debito  = carga[carga['DEBE HABER'] == 'D']['IMPORT_TOTAL'].sum()
    credito = carga[carga['DEBE HABER'] == 'H']['IMPORT_TOTAL'].sum()
    assert abs(debito - credito) <= 0.01
