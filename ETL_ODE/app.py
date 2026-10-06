from __future__ import annotations

import os
import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import pandas as pd

from transformer import (
    DEFAULT_CONFIG,
    get_app_dir,
    get_bank_default_config,
    load_config,
    normalize_bank_name,
    process_report,
    read_raw_ode_report,
    save_config,
    split_ode_rows,
    validate_report,
)

APP_DIR = get_app_dir()

_WARN_BG = '#fffbe6'
_WARN_FG = '#b35c00'
_BG_MAIN = '#edf4fb'
_BG_WHITE = '#ffffff'


def analyze_source(source_path: str) -> tuple:
    """Testable helper: returns (raw, excluded, processed) DataFrames."""
    raw = read_raw_ode_report(source_path)
    raw, excluded, processed = split_ode_rows(raw)
    return raw, excluded, processed


class ConversorODEApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title('Conversor de cobranzas ODE')
        self.geometry('1280x820')
        self.minsize(960, 600)
        self.configure(bg=_BG_MAIN)

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('TCombobox', fieldbackground='#f6fbff', foreground='#1a2c3f')

        self.file_label = tk.StringVar(value='')
        self.model_label = tk.StringVar(value='')
        self.output_dir = tk.StringVar(value='')
        self.year_var = tk.StringVar(value='2026')
        self.month_var = tk.StringVar(value='08')
        self.date_selector = tk.StringVar(value='fecha_cancelacion')
        self.config = load_config(APP_DIR / 'config.json')
        self.analysis_df = pd.DataFrame()
        self.last_result: dict = {}

        self._build_ui()
        self._apply_config_to_fields()

    # ------------------------------------------------------------------
    # Layout skeleton: fixed header ─ scrollable middle ─ fixed footer
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ── Fixed top banner ──────────────────────────────────────────
        self._build_header()

        # ── Fixed bottom action bar ───────────────────────────────────
        self._build_footer()

        # ── Scrollable middle canvas ──────────────────────────────────
        self._build_scroll_area()

        # Content sections go inside self._inner (the scrollable frame)
        self._build_file_selectors()
        self._build_config_section()
        self._build_banks_section()
        self._build_preview_section()
        self._build_validation_section()
        self._build_summary_cards()

        self.output_dir.set(str(APP_DIR / 'salida'))
        self.source_entry.focus_set()

    def _build_header(self) -> None:
        top = tk.Frame(self, bg='#0b1b2b')
        top.pack(side='top', fill='x')
        tk.Label(top, text='Conversor de cobranzas ODE',
                 bg='#0b1b2b', fg='#ffffff', font=('Segoe UI', 22, 'bold')
                 ).pack(anchor='w', padx=24, pady=(14, 2))
        tk.Label(top, text='Reporte de comprobantes cancelados \u2192 Formato de carga contable',
                 bg='#0b1b2b', fg='#7ec8ff', font=('Segoe UI', 10, 'bold')
                 ).pack(anchor='w', padx=24, pady=(0, 10))

    def _build_footer(self) -> None:
        """Action bar pinned to the bottom — always visible regardless of scroll position."""
        footer = tk.Frame(self, bg='#dce8f5', bd=0)
        footer.pack(side='bottom', fill='x')

        # separator line
        tk.Frame(footer, bg='#a8c4e0', height=1).pack(fill='x')

        bar = tk.Frame(footer, bg='#dce8f5', padx=14, pady=8)
        bar.pack(fill='x')

        self.progress = ttk.Progressbar(bar, orient='horizontal', mode='determinate', length=300)
        self.progress.pack(side='left', padx=(0, 10))
        self.status_label = tk.Label(bar, text='Listo', bg='#dce8f5', fg='#204a6d',
                                     font=('Segoe UI', 10, 'bold'))
        self.status_label.pack(side='left', padx=(0, 14))

        btn_defs = [
            ('Analizar archivos',       self.analyze_files,      '#1d4d82', 'white'),
            ('Generar carga',           self.generate_load,      '#2a9d8f', 'white'),
            ('Abrir carpeta de salida', self.open_output_folder, '#3a6ea5', 'white'),
            ('Restablecer',             self.reset_all,          '#cbcdd2', '#102a4a'),
            ('Guardar configuración',   self.save_config_fields, '#e0b03d', '#102a4a'),
        ]
        self.action_buttons: dict[str, tk.Button] = {}
        for text, cmd, bg, fg in btn_defs:
            btn = tk.Button(bar, text=text, command=cmd, bg=bg, fg=fg,
                            font=('Segoe UI', 10, 'bold'), padx=14, pady=6, relief='flat')
            btn.pack(side='left', padx=4)
            self.action_buttons[text] = btn

    def _build_scroll_area(self) -> None:
        """Canvas + scrollbar between header and footer."""
        wrapper = tk.Frame(self, bg=_BG_MAIN)
        wrapper.pack(side='top', fill='both', expand=True)

        vbar = ttk.Scrollbar(wrapper, orient='vertical')
        vbar.pack(side='right', fill='y')

        self._canvas = tk.Canvas(wrapper, bg=_BG_MAIN, highlightthickness=0,
                                 yscrollcommand=vbar.set)
        self._canvas.pack(side='left', fill='both', expand=True)
        vbar.configure(command=self._canvas.yview)

        self._inner = tk.Frame(self._canvas, bg=_BG_MAIN, padx=14, pady=10)
        self._inner_id = self._canvas.create_window((0, 0), window=self._inner, anchor='nw')

        # Keep inner frame width = canvas width
        self._canvas.bind('<Configure>', self._on_canvas_resize)
        self._inner.bind('<Configure>', self._on_inner_resize)

        # Mouse-wheel scrolling (Windows + Linux)
        self._canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        self._canvas.bind_all('<Button-4>', self._on_mousewheel)
        self._canvas.bind_all('<Button-5>', self._on_mousewheel)

    def _on_canvas_resize(self, event: tk.Event) -> None:
        self._canvas.itemconfig(self._inner_id, width=event.width)

    def _on_inner_resize(self, event: tk.Event) -> None:
        self._canvas.configure(scrollregion=self._canvas.bbox('all'))

    def _on_mousewheel(self, event: tk.Event) -> None:
        if event.num == 4:
            self._canvas.yview_scroll(-1, 'units')
        elif event.num == 5:
            self._canvas.yview_scroll(1, 'units')
        else:
            self._canvas.yview_scroll(int(-event.delta / 120), 'units')

    # ------------------------------------------------------------------
    # Content sections (all packed into self._inner)
    # ------------------------------------------------------------------

    def _section(self, title: str) -> tk.LabelFrame:
        return tk.LabelFrame(self._inner, text=title, bg=_BG_WHITE, fg='#143a57',
                             font=('Segoe UI', 11, 'bold'), padx=12, pady=8)

    def _build_file_selectors(self) -> None:
        frm = self._section('Archivos y periodo')
        frm.pack(fill='x', pady=(0, 10))

        for label_text, var, attr, cmd in [
            ('Reporte de cancelados',   self.file_label,  'source_entry', self.select_source),
            ('Excel modelo (opcional)', self.model_label, 'model_entry',  self.select_model),
            ('Carpeta de salida',       self.output_dir,  'output_entry', self.select_output),
        ]:
            row = tk.Frame(frm, bg=_BG_WHITE)
            row.pack(fill='x', pady=3)
            tk.Label(row, text=label_text, bg=_BG_WHITE, fg='#0d2140',
                     font=('Segoe UI', 10, 'bold'), width=22, anchor='w'
                     ).pack(side='left', padx=(0, 8))
            entry = tk.Entry(row, width=60, textvariable=var,
                             font=('Segoe UI', 10), relief='solid')
            entry.pack(side='left', fill='x', expand=True)
            tk.Button(row, text='Seleccionar', command=cmd,
                      bg='#1d4d82', fg='white', relief='flat', padx=8, pady=4
                      ).pack(side='left', padx=(6, 0))
            setattr(self, attr, entry)

        row4 = tk.Frame(frm, bg=_BG_WHITE)
        row4.pack(fill='x', pady=(6, 0))
        tk.Label(row4, text='Periodo contable', bg=_BG_WHITE, fg='#0d2140',
                 font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        tk.Label(row4, text='Año', bg=_BG_WHITE, fg='#0d2140').pack(side='left', padx=(0, 4))
        ttk.Combobox(row4, textvariable=self.year_var,
                     values=[str(y) for y in range(2024, 2036)],
                     width=7, state='readonly').pack(side='left', padx=(0, 14))
        tk.Label(row4, text='Mes', bg=_BG_WHITE, fg='#0d2140').pack(side='left', padx=(0, 4))
        ttk.Combobox(row4, textvariable=self.month_var,
                     values=[f'{i:02d}' for i in range(1, 13)],
                     width=5, state='readonly').pack(side='left', padx=(0, 14))
        tk.Label(row4, text='Fecha registro', bg=_BG_WHITE, fg='#0d2140').pack(side='left', padx=(0, 4))
        ttk.Combobox(row4, textvariable=self.date_selector,
                     values=['fecha_cancelacion', 'fecha_emi', 'ultimo_dia_mes'],
                     width=18, state='readonly').pack(side='left')

    def _build_config_section(self) -> None:
        frm = self._section('Configuración contable')
        frm.pack(fill='x', pady=(0, 10))
        self.config_entries: dict[str, tk.Entry] = {}
        for label, key in [
            ('Cuenta clientes', 'cuenta_clientes'),
            ('Sub diario',      'sub_diario'),
            ('Tipo anexo',      'tipo_anexo'),
            ('Tipo documento',  'tipo_documento'),
            ('Moneda',          'moneda'),
            ('Tipo conversión', 'tipo_conversion'),
            ('Doc. anulado',    'doc_anulado'),
            ('Glosa',           'glosa'),
        ]:
            row = tk.Frame(frm, bg=_BG_WHITE)
            row.pack(fill='x', pady=2)
            tk.Label(row, text=label, bg=_BG_WHITE, fg='#0d2140',
                     font=('Segoe UI', 10, 'bold'), width=18, anchor='w').pack(side='left')
            entry = tk.Entry(row, width=28, font=('Segoe UI', 10), relief='solid')
            entry.pack(side='left')
            self.config_entries[key] = entry

    def _build_banks_section(self) -> None:
        self.bank_frame = self._section('Bancos')
        self.bank_frame.pack(fill='x', pady=(0, 10))
        self.bank_entries: dict[str, dict] = {}
        for name in ['INTERBANK', 'BBVA CONTINENTAL', 'DE LA NACION', 'OTROS']:
            self._add_bank_row(name)

    def _build_preview_section(self) -> None:
        frm = self._section('Vista previa')
        frm.pack(fill='x', pady=(0, 10))
        self.preview = ttk.Treeview(
            frm,
            columns=('banco', 'serie', 'documento', 'importe', 'fecha'),
            show='headings', height=6)  # reduced height
        for col, w, anc in [
            ('banco', 150, 'w'), ('serie', 80, 'w'), ('documento', 100, 'w'),
            ('importe', 100, 'e'), ('fecha', 110, 'w'),
        ]:
            self.preview.heading(col, text=col.capitalize())
            self.preview.column(col, width=w, anchor=anc)
        sv = ttk.Scrollbar(frm, orient='vertical', command=self.preview.yview)
        self.preview.configure(yscrollcommand=sv.set)
        self.preview.pack(side='left', fill='x', expand=True)
        sv.pack(side='right', fill='y')

    def _build_validation_section(self) -> None:
        frm = self._section('Validaciones y advertencias')
        frm.pack(fill='x', pady=(0, 10))
        self.validation_box = tk.Text(frm, height=5, width=80, wrap='word',
                                      bg='#fffdf5', fg='#1b2a39', font=('Segoe UI', 10))
        self.validation_box.pack(fill='x', expand=True)

    def _build_summary_cards(self) -> None:
        cards_frame = tk.Frame(self._inner, bg=_BG_MAIN)
        cards_frame.pack(fill='x', pady=(0, 6))
        self.cards: dict[str, tk.Label] = {}
        for idx, (label, key) in enumerate([
            ('Registros encontrados', 'found'),
            ('Registros excluidos',   'excluded'),
            ('Registros procesados',  'processed'),
            ('Total cobrado',         'total'),
            ('Diferencia debe/haber', 'balance'),
            ('Estado final',          'status'),
        ]):
            card = tk.Frame(cards_frame, bg=_BG_WHITE,
                            highlightbackground='#dfeaf5', highlightthickness=1,
                            padx=10, pady=8)
            card.grid(row=0, column=idx, padx=6, pady=2, sticky='nsew')
            cards_frame.columnconfigure(idx, weight=1)
            tk.Label(card, text=label, bg=_BG_WHITE, fg='#56708a',
                     font=('Segoe UI', 8, 'bold')).pack(anchor='w')
            val = tk.Label(card, text='0', bg=_BG_WHITE, fg='#102a4a',
                           font=('Segoe UI', 16, 'bold'))
            val.pack(anchor='w', pady=(4, 0))
            self.cards[key] = val

    # ------------------------------------------------------------------
    # Bank row helper
    # ------------------------------------------------------------------

    def _add_bank_row(self, bank_name: str, warn: bool = False) -> None:
        if bank_name in self.bank_entries:
            return
        bg = _WARN_BG if warn else _BG_WHITE
        fg = _WARN_FG if warn else '#0d2140'
        row = tk.Frame(self.bank_frame, bg=bg)
        row.pack(fill='x', pady=2)
        tk.Label(row, text=bank_name, bg=bg, fg=fg,
                 font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        account = tk.Entry(row, width=16, font=('Segoe UI', 10), relief='solid')
        account.pack(side='left', padx=(0, 10))
        medio = tk.Entry(row, width=10, font=('Segoe UI', 10), relief='solid')
        medio.pack(side='left')
        self.bank_entries[bank_name] = {'cuenta': account, 'medio': medio}
        bank_cfg = self.config.get('bancos', {}).get(bank_name, get_bank_default_config(bank_name))
        account.insert(0, str(bank_cfg.get('cuenta_contable', get_bank_default_config(bank_name)['cuenta_contable'])))
        medio.insert(0, '' if bank_cfg.get('medio_pago') is None else str(bank_cfg.get('medio_pago', '')))

    # ------------------------------------------------------------------
    # Config helpers
    # ------------------------------------------------------------------

    def _apply_config_to_fields(self) -> None:
        for key, entry in self.config_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, str(self.config.get(key, DEFAULT_CONFIG.get(key, ''))))
        for bank_name, fields in self.bank_entries.items():
            bank_cfg = self.config.get('bancos', {}).get(bank_name, get_bank_default_config(bank_name))
            fields['cuenta'].delete(0, tk.END)
            fields['cuenta'].insert(0, str(bank_cfg.get('cuenta_contable', get_bank_default_config(bank_name)['cuenta_contable'])))
            fields['medio'].delete(0, tk.END)
            fields['medio'].insert(0, '' if bank_cfg.get('medio_pago') is None else str(bank_cfg.get('medio_pago', '')))

    def _collect_config(self) -> dict:
        cfg = self.config.copy()
        for key, entry in self.config_entries.items():
            cfg[key] = entry.get().strip()
        cfg['bancos'] = dict(cfg.get('bancos', {}))
        for bank_name, fields in self.bank_entries.items():
            defaults = get_bank_default_config(bank_name)
            cfg['bancos'][bank_name] = {
                'cuenta_contable': fields['cuenta'].get().strip() or defaults['cuenta_contable'],
                'medio_pago': fields['medio'].get().strip(),
            }
        return cfg

    def save_config_fields(self) -> None:
        self.config = self._collect_config()
        save_config(APP_DIR / 'config.json', self.config)
        messagebox.showinfo('Configuración guardada', 'La configuración contable se guardó correctamente.')

    # ------------------------------------------------------------------
    # File selectors
    # ------------------------------------------------------------------

    def select_source(self) -> None:
        path = filedialog.askopenfilename(
            title='Seleccionar reporte de cancelados',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')])
        if path:
            self.file_label.set(path)

    def select_model(self) -> None:
        path = filedialog.askopenfilename(
            title='Seleccionar modelo contable',
            filetypes=[('Excel', '*.xls *.xlsx'), ('Todos', '*.*')])
        if path:
            self.model_label.set(path)

    def select_output(self) -> None:
        folder = filedialog.askdirectory(title='Seleccionar carpeta de salida')
        if folder:
            self.output_dir.set(folder)

    # ------------------------------------------------------------------
    # Analyze
    # ------------------------------------------------------------------

    def analyze_files(self) -> None:
        source_path = self.file_label.get()
        if not source_path:
            messagebox.showwarning('Archivo requerido', 'Debe seleccionar un reporte de cancelados.')
            return
        periodo = f"{int(self.year_var.get()):04d}{self.month_var.get()}"
        fecha_calculo = self.date_selector.get()
        self.progress['value'] = 10
        self.status_label.config(text='Analizando reporte...')
        threading.Thread(
            target=self._analyze_worker,
            args=(source_path, periodo, fecha_calculo),
            daemon=True).start()

    def _analyze_worker(self, source_path: str, periodo: str, fecha_calculo: str) -> None:
        try:
            raw, excluded, processed = analyze_source(source_path)
            errors, warnings, info = validate_report(processed, periodo=periodo, fecha_calculo=fecha_calculo)

            new_banks: list[str] = []
            if 'banco' in processed.columns:
                for b in processed['banco'].astype(str).unique():
                    norm = normalize_bank_name(b)
                    if norm and norm not in self.bank_entries:
                        new_banks.append(norm)

            self.analysis_df = processed
            self.last_result = {
                'raw': raw, 'processed': processed, 'excluded': excluded,
                'errors': errors, 'warnings': warnings, 'info': info,
            }
            self.after(0, lambda: self._refresh_analysis(
                raw, excluded, processed, errors, warnings, info, new_banks))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror('Error', str(exc)))
            self.after(0, lambda: self.status_label.config(text='ERROR AL ANALIZAR'))
        finally:
            self.after(0, lambda: self.progress.configure(value=100))

    def _refresh_analysis(self, raw: pd.DataFrame, excluded: pd.DataFrame,
                          processed: pd.DataFrame, errors: list,
                          warnings: list, info: list, new_banks: list) -> None:
        state_after_analysis = 'LISTO PARA GENERAR' if not errors else 'REVISAR'
        self.status_label.config(text=state_after_analysis)

        for bank_name in new_banks:
            defaults = get_bank_default_config(bank_name)
            current = self.config.get('bancos', {}).get(bank_name, {})
            is_generic = (
                str(current.get('cuenta_contable', defaults['cuenta_contable'])).strip() == defaults['cuenta_contable']
                and str(current.get('medio_pago', defaults['medio_pago'])).strip() == str(defaults['medio_pago'])
            )
            self._add_bank_row(bank_name, warn=is_generic)
        if new_banks:
            warnings = list(warnings) + [
                f'Banco nuevo: {b}. Verifique la cuenta contable.' for b in new_banks]

        if 'banco' in processed.columns and 'importe_cancelado' in processed.columns:
            pv = processed[['banco', 'serie', 'documento', 'importe_cancelado', 'fecha_cancelacion']].head(100).copy()
            pv.columns = ['banco', 'serie', 'documento', 'importe', 'fecha']
            pv['importe'] = pv['importe'].apply(lambda x: f'{float(x):,.2f}' if pd.notna(x) else '0.00')
            self.preview.delete(*self.preview.get_children())
            for _, row in pv.iterrows():
                self.preview.insert('', 'end', values=(
                    row['banco'], row['serie'], row['documento'], row['importe'], str(row['fecha'])))

        total = float(processed['importe_cancelado'].sum()) if 'importe_cancelado' in processed.columns else 0.0
        self.cards['found'].config(text=str(len(raw)))
        self.cards['excluded'].config(text=str(len(excluded)))
        self.cards['processed'].config(text=str(len(processed)))
        self.cards['total'].config(text=f'S/{total:,.2f}')
        self.cards['balance'].config(text='PENDIENTE')
        self.cards['status'].config(text=state_after_analysis)

        lines = (
            [f'[ERROR] {m}' for m in errors]
            + [f'[ADVERTENCIA] {m}' for m in warnings]
            + [f'[INFO] {m}' for m in info]
        )
        self.validation_box.delete('1.0', tk.END)
        self.validation_box.insert('1.0', '\n'.join(lines))

        # Scroll canvas back to top after refresh
        self._canvas.yview_moveto(0.0)

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------

    def generate_load(self) -> None:
        source_path = self.file_label.get()
        output_dir = self.output_dir.get() or str(APP_DIR / 'salida')
        if not source_path:
            messagebox.showwarning('Archivo requerido', 'Debe seleccionar un reporte antes de generar la carga.')
            return
        periodo = f"{int(self.year_var.get()):04d}{self.month_var.get()}"
        fecha_calculo = self.date_selector.get()
        self.progress['value'] = 5
        self.status_label.config(text='Generando carga...')
        cfg = self._collect_config()
        self.config = cfg
        threading.Thread(
            target=self._generate_worker,
            args=(source_path, output_dir, cfg, periodo, fecha_calculo),
            daemon=True).start()

    def _generate_worker(self, source_path: str, output_dir: str, cfg: dict,
                         periodo: str, fecha_calculo: str) -> None:
        try:
            result = process_report(
                source_path, output_dir, cfg,
                fecha_calculo=fecha_calculo,
                periodo=periodo,
                model_path=self.model_label.get() or None,
            )
            self.after(0, lambda: self._finish_generation(result))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror('Error al generar', str(exc)))
            self.after(0, lambda: self.status_label.config(text='Error de generación'))
        finally:
            self.after(0, lambda: self.progress.configure(value=100))

    def _finish_generation(self, result: dict) -> None:
        self.last_result = result
        s = result['summary']
        diff = float(s.get('diferencia', 0.0))
        state_text = str(s.get('resultado', 'LISTO'))
        self.cards['found'].config(text=str(s.get('cantidad_original', 0)))
        self.cards['excluded'].config(text=str(s.get('cantidad_excluida', 0)))
        self.cards['processed'].config(text=str(s.get('cantidad_procesada', 0)))
        self.cards['total'].config(text=f"S/{float(s.get('total_procesado', 0.0)):,.2f}")
        self.cards['balance'].config(text=f"S/{diff:,.2f}")
        self.cards['status'].config(text=state_text)
        self.status_label.config(text=state_text)
        self.validation_box.delete('1.0', tk.END)
        lines = (
            [f'[ADVERTENCIA] {m}' for m in result.get('warnings', [])]
            + [f'[INFO] {m}' for m in result.get('info', [])]
        )
        self.validation_box.insert('1.0', '\n'.join(lines))
        messagebox.showinfo(
            'Carga generada',
            f'Archivo de carga (.xls):\n{result["output_file"]}\n\n'
            f'Pasos en la macro del estudio:\n'
            f'1. Usar "Pasar ventas y clientes a formato TXT"\n'
            f'   y seleccionar el archivo .xls anterior.\n'
            f'2. Cargar el TXT generado en el módulo contable.\n\n'
            f'Reporte de validación (solo revisión):\n{result["report_file"]}')

    # ------------------------------------------------------------------
    # Misc
    # ------------------------------------------------------------------

    def open_output_folder(self) -> None:
        folder = self.output_dir.get() or str(APP_DIR / 'salida')
        Path(folder).mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(folder)
        except Exception:
            webbrowser.open(folder)

    def reset_all(self) -> None:
        self.file_label.set('')
        self.model_label.set('')
        self.output_dir.set(str(APP_DIR / 'salida'))
        self.year_var.set('2026')
        self.month_var.set('08')
        self.date_selector.set('fecha_cancelacion')
        self.validation_box.delete('1.0', tk.END)
        self.preview.delete(*self.preview.get_children())
        self.progress['value'] = 0
        self.status_label.config(text='LISTO')
        self.config = load_config(APP_DIR / 'config.json')
        self._apply_config_to_fields()


if __name__ == '__main__':
    app = ConversorODEApp()
    app.mainloop()
