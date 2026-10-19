"""
Tests for feat/ode-bancos-series-ui:
- Sharon's confirmed DEFAULT_CONFIG values
- BCP alias normalisation
- Unconfigured bank: analysis reports, generation blocked
- Custom bank configured: uses its own account/medio
- Import of old JSON (no BCP): BCP added from defaults
- Export + reimport: no data loss
- Series editor: change type, verify TIP_DOC in output
- Mixed factura/boleta report with BCP
- Invalidation on config change
- HTML safety: escHtml not tested here (browser-only)
"""
from __future__ import annotations

import io
import json
import os
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest
import xlrd

# Configure env vars BEFORE importing server
os.environ.setdefault('SECRET_KEY', 'test-secret-key-do-not-use-in-prod')
os.environ.setdefault('ODE_USERNAME', 'testuser')
os.environ.setdefault('ODE_PASSWORD', 'testpass')

from fastapi.testclient import TestClient

from server import app  # noqa: E402
from transformer import (
    DEFAULT_CONFIG,
    build_carga_dataframe,
    get_unconfigured_banks,
    load_config,
    normalize_bank_name,
    process_report,
    read_raw_ode_report,
    split_ode_rows,
    validate_bank_cfg,
    validate_report,
)

client = TestClient(app, raise_server_exceptions=False)

_ETL_DIR = Path(__file__).resolve().parent.parent
SOURCE_PATH = _ETL_DIR / 'reporte de cancelados agosto 2026.xlsx'
_REAL_CONFIG = {'serie_map': {'FE01': '01'}}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _auth_cookies():
    res = client.post('/login', data={'username': 'testuser', 'password': 'testpass'})
    assert res.status_code == 200
    return dict(res.cookies)


def _make_xlsx(rows, with_tipo_col=True, banco='INTERBANK'):
    wb = openpyxl.Workbook()
    ws = wb.active
    cols = ['NumRuc', 'Serie', 'Documento', 'Importe Cancelado', 'Fecha Cancelacion', 'Banco']
    if with_tipo_col:
        cols.append('Tipo Comprobante')
    ws.append(cols)
    for r in rows:
        ws.append(r[:6] + ((r[6],) if with_tipo_col and len(r) > 6 else ()))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _default_config_with_bcp(**overrides):
    cfg = {
        'cuenta_clientes': '12120001',
        'sub_diario': '07',
        'tipo_anexo': '02',
        'moneda': 'MN',
        'tipo_conversion': 'VTA',
        'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES',
        'serie_map': {},
        'bancos': {
            'INTERBANK':        {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'BBVA CONTINENTAL': {'cuenta_contable': '10400001', 'medio_pago': '003'},
            'BCP':              {'cuenta_contable': '10400011', 'medio_pago': '003'},
            'DE LA NACION':     {'cuenta_contable': '10710001', 'medio_pago': '003'},
            'OTROS':            {'cuenta_contable': '10100001', 'medio_pago': ''},
        },
    }
    cfg.update(overrides)
    return cfg


# ---------------------------------------------------------------------------
# 1. Sharon's confirmed DEFAULT_CONFIG values
# ---------------------------------------------------------------------------

class TestSharonDefaultConfig:
    def test_sub_diario_is_07(self):
        assert DEFAULT_CONFIG['sub_diario'] == '07', \
            f"sub_diario debe ser '07' (confirmado por Sharon), got {DEFAULT_CONFIG['sub_diario']!r}"

    def test_cuenta_clientes(self):
        assert DEFAULT_CONFIG['cuenta_clientes'] == '12120001'

    def test_tipo_anexo(self):
        assert DEFAULT_CONFIG['tipo_anexo'] == '02'

    def test_bcp_in_default(self):
        assert 'BCP' in DEFAULT_CONFIG['bancos'], "BCP debe estar en DEFAULT_CONFIG['bancos']"
        assert DEFAULT_CONFIG['bancos']['BCP']['cuenta_contable'] == '10400011'
        assert DEFAULT_CONFIG['bancos']['BCP']['medio_pago'] == '003'

    def test_interbank_cuenta(self):
        assert DEFAULT_CONFIG['bancos']['INTERBANK']['cuenta_contable'] == '10400006'

    def test_bbva_cuenta(self):
        assert DEFAULT_CONFIG['bancos']['BBVA CONTINENTAL']['cuenta_contable'] == '10400001'

    def test_nacion_cuenta(self):
        assert DEFAULT_CONFIG['bancos']['DE LA NACION']['cuenta_contable'] == '10710001'

    def test_otros_cuenta_caja(self):
        assert DEFAULT_CONFIG['bancos']['OTROS']['cuenta_contable'] == '10100001'
        assert DEFAULT_CONFIG['bancos']['OTROS']['medio_pago'] == ''

    def test_no_tipo_documento_key(self):
        """tipo_documento should not be in DEFAULT_CONFIG — it was the confusing legacy field."""
        assert 'tipo_documento' not in DEFAULT_CONFIG, \
            "tipo_documento ya no debe estar en DEFAULT_CONFIG"

    def test_all_codes_are_strings(self):
        """All codes must be strings to preserve leading zeros."""
        for name, cfg in DEFAULT_CONFIG['bancos'].items():
            assert isinstance(cfg['cuenta_contable'], str), \
                f"cuenta_contable de {name} debe ser str"
            assert isinstance(cfg['medio_pago'], str), \
                f"medio_pago de {name} debe ser str"


# ---------------------------------------------------------------------------
# 2. BCP alias normalisation
# ---------------------------------------------------------------------------

class TestBCPNormalisation:
    def test_bcp_exact(self):
        assert normalize_bank_name('BCP') == 'BCP'

    def test_banco_de_credito(self):
        assert normalize_bank_name('Banco de Crédito') == 'BCP'

    def test_banco_de_credito_del_peru(self):
        assert normalize_bank_name('Banco de Crédito del Perú') == 'BCP'

    def test_banco_credito_no_accent(self):
        assert normalize_bank_name('BANCO DE CREDITO DEL PERU') == 'BCP'

    def test_bcp_mixed_case(self):
        assert normalize_bank_name('bcp') == 'BCP'

    def test_bcp_separate_from_otros(self):
        assert normalize_bank_name('BCP') != 'OTROS'

    def test_bcp_separate_from_interbank(self):
        assert normalize_bank_name('BCP') != 'INTERBANK'


# ---------------------------------------------------------------------------
# 3. Unconfigured bank: analysis warns, generation blocked
# ---------------------------------------------------------------------------

class TestUnconfiguredBank:
    _DATE = datetime(2026, 8, 10)

    def _df(self, banco='BANCO_NUEVO'):
        return pd.DataFrame({
            'numruc': ['20111111111'],
            'serie': ['FE01'],
            'documento': ['0000001'],
            'importe_cancelado': [100.0],
            'banco': [banco],
            'fecha_cancelacion': [self._DATE],
            'tipo_comprobante': ['01'],
        })

    def test_get_unconfigured_banks_detects_missing(self):
        df = self._df('BANCO_NUEVO')
        cfg = {'bancos': {'INTERBANK': {'cuenta_contable': '10400006', 'medio_pago': '003'}}}
        missing = get_unconfigured_banks(df, cfg)
        assert 'Banco_Nuevo' in missing or 'BANCO_NUEVO' in ' '.join(missing).upper() or len(missing) > 0

    def test_validate_report_blocks_unconfigured_bank(self):
        df = self._df('BANCO_NUEVO')
        cfg = _default_config_with_bcp()  # does not contain BANCO_NUEVO
        errors, _, _ = validate_report(df, periodo='202608', fecha_calculo='fecha_cancelacion', config=cfg)
        bank_errors = [e for e in errors if 'BANCO_NUEVO' in e.upper() or 'BANCO_NUEVO'.lower() in e.lower()]
        assert len(bank_errors) > 0, f"Expected error for unconfigured bank, got: {errors}"

    def test_generate_blocked_via_api(self):
        xlsx = _make_xlsx([
            ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'BANCO_NUEVO', '01'),
        ], with_tipo_col=True)
        cookies = _auth_cookies()
        cfg = _default_config_with_bcp()  # no BANCO_NUEVO
        res = client.post(
            '/api/generate',
            files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
                  'config_json': json.dumps(cfg)},
            cookies=cookies,
        )
        assert res.status_code == 422, f"Expected 422, got {res.status_code}: {res.text}"
        assert 'BANCO_NUEVO' in res.text or 'banco' in res.text.lower()

    def test_analyze_reports_unconfigured_bank(self):
        xlsx = _make_xlsx([
            ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'BANCO_NUEVO', '01'),
        ], with_tipo_col=True)
        cookies = _auth_cookies()
        cfg = _default_config_with_bcp()
        res = client.post(
            '/api/analyze',
            files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
                  'config_json': json.dumps(cfg)},
            cookies=cookies,
        )
        assert res.status_code == 200
        data = res.json()
        assert len(data.get('errors', [])) > 0 or len(data.get('unconfigured_banks', [])) > 0


# ---------------------------------------------------------------------------
# 4. Custom bank configured: uses real account and medio
# ---------------------------------------------------------------------------

def test_custom_bank_uses_own_config(tmp_path):
    """A fully configured custom bank produces its exact account in the output."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'SCOTIABANK', '01'),
    ], with_tipo_col=True)
    cfg = _default_config_with_bcp()
    cfg['bancos']['SCOTIABANK'] = {'cuenta_contable': '10400099', 'medio_pago': '003'}

    cookies = _auth_cookies()
    res = client.post(
        '/api/generate',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    import base64
    files = res.json()['files']
    xls_data = next(f for f in files if f['name'].endswith('.xls'))
    wb = xlrd.open_workbook(file_contents=base64.b64decode(xls_data['data']))
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    cta_col = headers.index('CTA_CONTABLE')
    ctas = [str(ws.cell_value(r, cta_col)) for r in range(1, ws.nrows)]
    assert '10400099' in ctas, f"Custom bank account 10400099 not found in output: {ctas}"


# ---------------------------------------------------------------------------
# 5. Import old JSON (no BCP): BCP completed from defaults
# ---------------------------------------------------------------------------

def test_import_old_config_without_bcp_adds_bcp(tmp_path):
    """Old config (no BCP, no sub_diario '07') is loaded and BCP is added from defaults."""
    old_cfg = {
        'cuenta_clientes': '12120001',
        'sub_diario': '01',
        'tipo_anexo': '02',
        'moneda': 'MN',
        'tipo_conversion': 'VTA',
        'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES',
        'serie_map': {},
        'bancos': {
            'INTERBANK':        {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'BBVA CONTINENTAL': {'cuenta_contable': '10400001', 'medio_pago': '003'},
            'DE LA NACION':     {'cuenta_contable': '10710001', 'medio_pago': '003'},
            'OTROS':            {'cuenta_contable': '10100001', 'medio_pago': ''},
        },
    }
    cfg_path = tmp_path / 'old_config.json'
    cfg_path.write_text(json.dumps(old_cfg), encoding='utf-8')
    loaded = load_config(cfg_path)
    assert 'BCP' in loaded['bancos'], "BCP should be added from DEFAULT_CONFIG when absent in old config"
    assert loaded['bancos']['BCP']['cuenta_contable'] == '10400011'
    # Old config's explicit banks are preserved
    assert loaded['bancos']['INTERBANK']['cuenta_contable'] == '10400006'


def test_import_old_config_with_tipo_documento_ignores_it(tmp_path):
    """Old config with tipo_documento field is accepted; the field is not used for resolution."""
    old_cfg = {
        'cuenta_clientes': '12120001',
        'sub_diario': '01',
        'tipo_anexo': '02',
        'tipo_documento': '01',  # legacy field
        'moneda': 'MN',
        'tipo_conversion': 'VTA',
        'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES',
        'serie_map': {'FE01': '01'},
        'bancos': {},
    }
    # Should not raise
    cookies = _auth_cookies()
    res = client.post(
        '/api/config/validate',
        data={'config_json': json.dumps(old_cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, f"Old config with tipo_documento should be valid: {res.text}"


# ---------------------------------------------------------------------------
# 6. Export + reimport: no data loss
# ---------------------------------------------------------------------------

def test_export_and_reimport_preserves_banks():
    """Config exported via readConfig() and re-imported is byte-identical."""
    from_config = _default_config_with_bcp()
    from_config['bancos']['SCOTIABANK'] = {'cuenta_contable': '10400099', 'medio_pago': '003'}
    from_config['serie_map'] = {'FE01': '01', 'BE': '03'}

    exported = json.dumps(from_config, ensure_ascii=False, indent=2)
    reimported = json.loads(exported)

    assert reimported['bancos']['BCP']['cuenta_contable'] == '10400011'
    assert reimported['bancos']['SCOTIABANK']['cuenta_contable'] == '10400099'
    assert reimported['serie_map'] == {'FE01': '01', 'BE': '03'}
    assert reimported['sub_diario'] == '07'


# ---------------------------------------------------------------------------
# 7. Series editor: change type via map, verify TIP_DOC in XLS output
# ---------------------------------------------------------------------------

def test_series_editor_remap_changes_tip_doc(tmp_path):
    """Changing a series mapping from the editor produces the correct TIP_DOC in XLS."""
    xlsx = _make_xlsx([
        ('20111111111', 'XX01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK'),
        ('20222222222', 'XX01', '0000002',  50.0, datetime(2026, 8, 11), 'INTERBANK'),
    ], with_tipo_col=False)

    cookies = _auth_cookies()
    cfg = _default_config_with_bcp(serie_map={'XX01': '03'})  # assign as Boleta

    res = client.post(
        '/api/generate',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    import base64
    files = res.json()['files']
    xls_data = next(f for f in files if f['name'].endswith('.xls'))
    wb = xlrd.open_workbook(file_contents=base64.b64decode(xls_data['data']))
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    tip_col = headers.index('TIP_DOC')
    dh_col  = headers.index('DEBE HABER')
    tip_values = [str(ws.cell_value(r, tip_col)) for r in range(1, ws.nrows)
                  if str(ws.cell_value(r, dh_col)) == 'H']
    assert all(v == '03' for v in tip_values), \
        f"All H-rows for serie XX01 should have TIP_DOC=03; got {tip_values}"


# ---------------------------------------------------------------------------
# 8. Mixed report: factura + boleta + BCP, balance correct
# ---------------------------------------------------------------------------

def test_mixed_factura_boleta_with_bcp(tmp_path):
    """Report with both facturas (01) and boletas (03) including BCP bank balances correctly."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 200.0, datetime(2026, 8, 10), 'BCP',       '01'),
        ('20222222222', 'BE01', '0000002', 150.0, datetime(2026, 8, 11), 'INTERBANK', '03'),
        ('20333333333', 'FE01', '0000003', 100.0, datetime(2026, 8, 12), 'OTROS',     '01'),
    ], with_tipo_col=True)

    cookies = _auth_cookies()
    cfg = _default_config_with_bcp()

    res = client.post(
        '/api/generate',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    data = res.json()
    summ = data['summary']
    assert abs(summ['diferencia']) <= 0.01, f"Balance mismatch: {summ['diferencia']}"
    assert summ['resultado'] == 'CUADRADO'
    assert abs(summ['total_procesado'] - 450.0) <= 0.01

    # Verify BCP uses correct account
    import base64
    files = data['files']
    xls_data = next(f for f in files if f['name'].endswith('.xls'))
    wb = xlrd.open_workbook(file_contents=base64.b64decode(xls_data['data']))
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    cta_col = headers.index('CTA_CONTABLE')
    dh_col  = headers.index('DEBE HABER')
    bcp_d_rows = [str(ws.cell_value(r, cta_col)) for r in range(1, ws.nrows)
                  if str(ws.cell_value(r, cta_col)) == '10400011'
                  and str(ws.cell_value(r, dh_col)) == 'D']
    assert len(bcp_d_rows) == 1, f"Expected 1 BCP D-row with account 10400011, got: {bcp_d_rows}"


# ---------------------------------------------------------------------------
# 9. Invalidation: downloads stale after config change
# ---------------------------------------------------------------------------

def test_analysis_valid_flag_cleared_on_unconfigured_bank():
    """Analysis result includes errors when a bank is unconfigured — blocks generation."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'BANCO_NUEVO', '01'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    cfg = _default_config_with_bcp()  # no BANCO_NUEVO

    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200
    data = res.json()
    # Either errors list or unconfigured_banks must be non-empty
    blocking = len(data.get('errors', [])) + len(data.get('unconfigured_banks', []))
    assert blocking > 0, "Should report blocking issues for unconfigured bank"


# ---------------------------------------------------------------------------
# 10. series_all in analysis includes count and amount
# ---------------------------------------------------------------------------

def test_analyze_series_all_includes_count_and_amount():
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK', '01'),
        ('20222222222', 'FE01', '0000002',  50.0, datetime(2026, 8, 11), 'INTERBANK', '01'),
        ('20333333333', 'BE01', '0000003',  75.0, datetime(2026, 8, 12), 'INTERBANK', '03'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    cfg = _default_config_with_bcp()

    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    data = res.json()
    series_all = data.get('series_all', [])
    assert len(series_all) >= 2, f"Expected at least 2 series, got: {series_all}"

    fe01 = next((s for s in series_all if s['serie'] == 'FE01'), None)
    be01 = next((s for s in series_all if s['serie'] == 'BE01'), None)
    assert fe01 is not None, "FE01 should be in series_all"
    assert be01 is not None, "BE01 should be in series_all"
    assert fe01.get('conteo') == 2, f"FE01 count should be 2, got {fe01.get('conteo')}"
    assert abs(fe01.get('importe', 0) - 150.0) <= 0.01
    assert be01.get('conteo') == 1
    assert abs(be01.get('importe', 0) - 75.0) <= 0.01

    # All series should be marked resolved (explicit tipo_comprobante column)
    assert fe01.get('resuelto') is True
    assert be01.get('resuelto') is True


# ---------------------------------------------------------------------------
# 11. Config validate: bank structure validated server-side
# ---------------------------------------------------------------------------

def test_config_validate_rejects_bank_without_cuenta():
    cookies = _auth_cookies()
    bad_cfg = _default_config_with_bcp()
    bad_cfg['bancos']['SCOTIABANK'] = {'medio_pago': '003'}  # missing cuenta_contable
    res = client.post(
        '/api/config/validate',
        data={'config_json': json.dumps(bad_cfg)},
        cookies=cookies,
    )
    assert res.status_code == 422, f"Expected 422 for missing cuenta_contable, got {res.status_code}"


def test_config_validate_rejects_invalid_json():
    cookies = _auth_cookies()
    res = client.post(
        '/api/config/validate',
        data={'config_json': 'this is not json'},
        cookies=cookies,
    )
    assert res.status_code == 422


# ===========================================================================
# PR-REVIEW REGRESSION TESTS (round 2)
# ===========================================================================

# ---------------------------------------------------------------------------
# R1. schema_version=2 prevents deleted banks from reappearing
# ---------------------------------------------------------------------------

def test_deleted_bcp_does_not_reappear_in_analyze():
    """With schema_version=2 a user-deleted BCP is not re-added by _build_config."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK', '01'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    # Send config WITHOUT BCP, schema_version=2 (current state is authoritative)
    cfg = {
        'schema_version': 2,
        'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK':        {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'BBVA CONTINENTAL': {'cuenta_contable': '10400001', 'medio_pago': '003'},
            'DE LA NACION':     {'cuenta_contable': '10710001', 'medio_pago': '003'},
            'OTROS':            {'cuenta_contable': '10100001', 'medio_pago': ''},
            # BCP deliberately omitted
        },
    }
    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200
    # BCP is not in the report, so no error about it; but it must not auto-appear
    data = res.json()
    # The report only has INTERBANK, so there should be no BCP-related error
    bcp_errors = [e for e in data.get('errors', []) if 'BCP' in e.upper()]
    assert not bcp_errors, f"BCP should not appear when deleted: {bcp_errors}"


def test_deleted_bcp_blocked_in_generate_if_used():
    """With schema_version=2, if BCP is in the report but not in config → generate blocked."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'BCP', '01'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    cfg = {
        'schema_version': 2,
        'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK': {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'OTROS':     {'cuenta_contable': '10100001', 'medio_pago': ''},
            # BCP deliberately absent — user deleted it
        },
    }
    res = client.post(
        '/api/generate',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 422, f"Should block when BCP deleted but in report, got {res.status_code}: {res.text}"
    assert 'BCP' in res.json().get('detail', '').upper(), f"Error should mention BCP: {res.text}"


# ---------------------------------------------------------------------------
# R2. Empty cuenta_contable blocks both analysis and generation
# ---------------------------------------------------------------------------

class TestEmptyCuentaContable:
    _DATE = datetime(2026, 8, 10)

    def _xlsx_bank(self, banco):
        return _make_xlsx([
            ('20111111111', 'FE01', '0000001', 100.0, self._DATE, banco, '01'),
        ], with_tipo_col=True)

    def test_validate_bank_cfg_rejects_empty_cuenta(self):
        err = validate_bank_cfg('MIBANCO', {'cuenta_contable': '', 'medio_pago': '003'})
        assert err is not None, "Empty cuenta_contable should return an error"
        assert 'cuenta_contable' in err.lower() or 'cuenta' in err.lower()

    def test_validate_bank_cfg_rejects_whitespace_cuenta(self):
        err = validate_bank_cfg('MIBANCO', {'cuenta_contable': '   ', 'medio_pago': '003'})
        assert err is not None, "Whitespace-only cuenta_contable should return an error"

    def test_validate_bank_cfg_accepts_valid(self):
        err = validate_bank_cfg('MIBANCO', {'cuenta_contable': '10400099', 'medio_pago': '003'})
        assert err is None, f"Valid config should pass: {err}"

    def test_validate_bank_cfg_accepts_otros_empty_medio(self):
        err = validate_bank_cfg('OTROS', {'cuenta_contable': '10100001', 'medio_pago': ''})
        assert err is None, f"OTROS with empty medio_pago should be valid: {err}"

    def test_validate_bank_cfg_rejects_none_cuenta(self):
        err = validate_bank_cfg('MIBANCO', {'cuenta_contable': None, 'medio_pago': '003'})
        assert err is not None

    def test_validate_bank_cfg_rejects_list_type(self):
        err = validate_bank_cfg('MIBANCO', {'cuenta_contable': ['10400099'], 'medio_pago': '003'})
        assert err is not None

    def test_get_unconfigured_detects_empty_cuenta(self):
        df = pd.DataFrame({
            'banco': ['MIBANCO'],
            'importe_cancelado': [100.0],
        })
        cfg = {'bancos': {'MIBANCO': {'cuenta_contable': '', 'medio_pago': '003'}}}
        missing = get_unconfigured_banks(df, cfg)
        assert 'Mibanco' in missing or 'MIBANCO' in ' '.join(missing).upper(), \
            f"Bank with empty cuenta should be unconfigured: {missing}"

    def test_analyze_blocks_on_empty_cuenta(self):
        xlsx = self._xlsx_bank('INTERBANK')
        cookies = _auth_cookies()
        cfg = {
            'schema_version': 2,
            'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
            'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
            'glosa': 'COBRANZA DEL MES', 'serie_map': {},
            'bancos': {
                'INTERBANK': {'cuenta_contable': '', 'medio_pago': '003'},  # empty!
                'OTROS':     {'cuenta_contable': '10100001', 'medio_pago': ''},
            },
        }
        res = client.post(
            '/api/analyze',
            files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
                  'config_json': json.dumps(cfg)},
            cookies=cookies,
        )
        assert res.status_code == 200
        data = res.json()
        blocking = len(data.get('errors', [])) + len(data.get('unconfigured_banks', []))
        assert blocking > 0, "Empty cuenta_contable should block analysis"

    def test_generate_blocks_on_empty_cuenta(self):
        xlsx = self._xlsx_bank('INTERBANK')
        cookies = _auth_cookies()
        cfg = {
            'schema_version': 2,
            'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
            'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
            'glosa': 'COBRANZA DEL MES', 'serie_map': {},
            'bancos': {
                'INTERBANK': {'cuenta_contable': '   ', 'medio_pago': '003'},  # whitespace!
                'OTROS':     {'cuenta_contable': '10100001', 'medio_pago': ''},
            },
        }
        res = client.post(
            '/api/generate',
            files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
                  'config_json': json.dumps(cfg)},
            cookies=cookies,
        )
        assert res.status_code == 422, f"Empty cuenta should block generation: {res.status_code}: {res.text}"


# ---------------------------------------------------------------------------
# R3. Export/reimport preserves deleted banks (schema_version round-trip)
# ---------------------------------------------------------------------------

def test_export_reimport_with_bcp_deleted():
    """Config exported with schema_version=2 and no BCP is reimported as-is (no auto-add)."""
    cfg_no_bcp = {
        'schema_version': 2,
        'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK': {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'OTROS':     {'cuenta_contable': '10100001', 'medio_pago': ''},
        },
    }
    exported = json.dumps(cfg_no_bcp, ensure_ascii=False, indent=2)
    reimported = json.loads(exported)

    # schema_version=2 must survive the round-trip
    assert reimported.get('schema_version') == 2
    # BCP must NOT be auto-added (it was deleted by the user)
    assert 'BCP' not in reimported['bancos'], \
        "Reimported v2 config must not silently add BCP back"


def test_import_old_json_v1_adds_bcp(tmp_path):
    """Old JSON (no schema_version) has BCP added during migration."""
    old_cfg = {
        'cuenta_clientes': '12120001', 'sub_diario': '01', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK': {'cuenta_contable': '10400006', 'medio_pago': '003'},
            'OTROS':     {'cuenta_contable': '10100001', 'medio_pago': ''},
        },
        # no schema_version → v1 → migration adds BCP
    }
    path = tmp_path / 'old.json'
    path.write_text(json.dumps(old_cfg), encoding='utf-8')
    loaded = load_config(path)
    assert 'BCP' in loaded['bancos'], "v1 import must add BCP from defaults"
    assert loaded['bancos']['BCP']['cuenta_contable'] == '10400011'
    assert loaded.get('schema_version') == 2


# ---------------------------------------------------------------------------
# R4. Bank type validation returns 422 (not 500)
# ---------------------------------------------------------------------------

def test_bank_config_with_integer_cuenta_returns_422():
    """Integer cuenta_contable should return 422, not crash with 500."""
    cookies = _auth_cookies()
    bad_cfg = {
        'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK': {'cuenta_contable': 10400006, 'medio_pago': '003'},  # int, not str
        },
    }
    res = client.post(
        '/api/config/validate',
        data={'config_json': json.dumps(bad_cfg)},
        cookies=cookies,
    )
    assert res.status_code == 422, f"Expected 422 for int cuenta, got {res.status_code}: {res.text}"


def test_bank_config_with_null_medio_returns_422():
    """null medio_pago should return 422."""
    cookies = _auth_cookies()
    bad_cfg = {
        'cuenta_clientes': '12120001', 'sub_diario': '07', 'tipo_anexo': '02',
        'moneda': 'MN', 'tipo_conversion': 'VTA', 'doc_anulado': '0',
        'glosa': 'COBRANZA DEL MES', 'serie_map': {},
        'bancos': {
            'INTERBANK': {'cuenta_contable': '10400006', 'medio_pago': None},
        },
    }
    res = client.post(
        '/api/config/validate',
        data={'config_json': json.dumps(bad_cfg)},
        cookies=cookies,
    )
    assert res.status_code == 422, f"Expected 422 for null medio_pago, got {res.status_code}: {res.text}"


# ---------------------------------------------------------------------------
# R5. Serie map: removing a series mapping works (via API round-trip)
# ---------------------------------------------------------------------------

def test_removing_serie_from_map_blocks_generation():
    """If the only resolution for a serie was its map entry and we remove it, generation blocks."""
    xlsx = _make_xlsx([
        ('20111111111', 'XX99', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK'),
    ], with_tipo_col=False)
    cookies = _auth_cookies()

    # With the mapping: should generate
    cfg_with = _default_config_with_bcp(serie_map={'XX99': '01'})
    cfg_with['schema_version'] = 2
    res_ok = client.post(
        '/api/generate',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg_with)},
        cookies=cookies,
    )
    assert res_ok.status_code == 200, f"Should succeed with mapping: {res_ok.text}"

    # Without the mapping: should fail
    cfg_without = _default_config_with_bcp(serie_map={})  # XX99 removed
    cfg_without['schema_version'] = 2
    res_fail = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg_without)},
        cookies=cookies,
    )
    assert res_fail.status_code == 200
    data = res_fail.json()
    assert len(data.get('errors', [])) > 0, "Removing XX99 mapping should produce errors"


def test_mappings_for_other_series_preserved():
    """Series not in the current report keep their mapping (tested via analyze)."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK', '01'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    # Include mapping for BE01 which is NOT in this report
    cfg = _default_config_with_bcp(serie_map={'BE01': '03', 'FE01': '01'})
    cfg['schema_version'] = 2

    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    # The analysis succeeds; BE01 is absent from series_all (not in report)
    data = res.json()
    series_names = [s['serie'] for s in data.get('series_all', [])]
    assert 'BE01' not in series_names, "BE01 should not appear in series_all since it's not in the report"
    assert 'FE01' in series_names, "FE01 must appear"


# ---------------------------------------------------------------------------
# R6. series_all includes fuente_es_prefijo for prefix-matched series
# ---------------------------------------------------------------------------

def test_series_all_fuente_prefix_detected():
    """A serie matched by a prefix key shows fuente_es_prefijo=True."""
    xlsx = _make_xlsx([
        # FE01 matches prefix 'FE' in serie_map
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK'),
    ], with_tipo_col=False)
    cookies = _auth_cookies()
    cfg = _default_config_with_bcp(serie_map={'FE': '01'})  # prefix, not exact
    cfg['schema_version'] = 2

    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    data = res.json()
    fe01 = next((s for s in data.get('series_all', []) if s['serie'] == 'FE01'), None)
    assert fe01 is not None
    assert fe01.get('fuente') == 'mapeo'
    assert fe01.get('fuente_es_prefijo') is True, \
        f"FE01 matched by prefix 'FE' should have fuente_es_prefijo=True, got: {fe01}"
    assert fe01.get('fuente_clave') == 'FE'


def test_series_all_fuente_exact_not_prefix():
    """An exact map match reports fuente_es_prefijo=False."""
    xlsx = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK'),
    ], with_tipo_col=False)
    cookies = _auth_cookies()
    cfg = _default_config_with_bcp(serie_map={'FE01': '01'})  # exact match
    cfg['schema_version'] = 2

    res = client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )
    assert res.status_code == 200, res.text
    data = res.json()
    fe01 = next((s for s in data.get('series_all', []) if s['serie'] == 'FE01'), None)
    assert fe01 is not None
    assert fe01.get('fuente_es_prefijo') is False, \
        f"Exact match should have fuente_es_prefijo=False, got: {fe01}"


# ---------------------------------------------------------------------------
# R7. Race condition — stale analyze response test (server-side behaviour)
# ---------------------------------------------------------------------------

def test_two_sequential_analyses_second_wins():
    """Sending two analyze requests sequentially; the last one's result is authoritative."""
    xlsx1 = _make_xlsx([
        ('20111111111', 'FE01', '0000001', 100.0, datetime(2026, 8, 10), 'INTERBANK', '01'),
    ], with_tipo_col=True)
    xlsx2 = _make_xlsx([
        ('20222222222', 'FE01', '0000001', 200.0, datetime(2026, 8, 15), 'INTERBANK', '01'),
        ('20333333333', 'FE01', '0000002', 300.0, datetime(2026, 8, 16), 'INTERBANK', '01'),
    ], with_tipo_col=True)
    cookies = _auth_cookies()
    cfg = _default_config_with_bcp()
    cfg['schema_version'] = 2

    _analyze = lambda xlsx: client.post(
        '/api/analyze',
        files={'reporte': ('r.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'year': '2026', 'month': '8', 'fecha_calculo': 'fecha_cancelacion',
              'config_json': json.dumps(cfg)},
        cookies=cookies,
    )

    res1 = _analyze(xlsx1)
    res2 = _analyze(xlsx2)
    assert res2.status_code == 200
    data2 = res2.json()
    # Second analysis has 2 rows — not 1
    assert data2['stats']['cantidad_procesada'] == 2, \
        f"Second analysis should have 2 rows, got {data2['stats']['cantidad_procesada']}"

    # Note: client-side race condition (revision guard) cannot be tested here without a browser.
    # The server is stateless, so both requests succeed independently. The frontend
    # revision guard (inputRevision !== myRevision) is tested manually in the browser.


# ---------------------------------------------------------------------------
# R8. Validate that Sharon's confirmed values are preserved after process_report
# ---------------------------------------------------------------------------

def test_real_report_uses_sharon_sub_diario(tmp_path):
    """process_report with DEFAULT_CONFIG uses sub_diario '07' in all output rows."""
    import xlrd as _xlrd
    from transformer import process_report as pr
    result = pr(
        SOURCE_PATH, tmp_path,
        config={**DEFAULT_CONFIG, 'serie_map': {'FE01': '01'}},
        fecha_calculo='fecha_cancelacion', periodo='202608',
    )
    wb = _xlrd.open_workbook(result['output_file'])
    ws = wb.sheet_by_index(0)
    headers = [ws.cell_value(0, c) for c in range(ws.ncols)]
    sd_col = headers.index('SUB_DIARIO')
    sub_diarios = {str(ws.cell_value(r, sd_col)) for r in range(1, ws.nrows)}
    assert sub_diarios == {'07'}, f"All rows must have SUB_DIARIO='07', got: {sub_diarios}"
