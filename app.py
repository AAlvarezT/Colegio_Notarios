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
    build_carga_dataframe,
    load_config,
    load_ode_report,
    process_report,
    save_config,
    validate_report,
)


class ConversorODEApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title('Conversor de cobranzas ODE')
        self.geometry('1280x840')
        self.minsize(1180, 760)
        self.configure(bg='#f4f7fb')

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('Dark.TFrame', background='#0b1b2b')
        self.style.configure('Card.TFrame', background='#ffffff')
        self.style.configure('Title.TLabel', background='#0b1b2b', foreground='#f4f7fb', font=('Segoe UI', 18, 'bold'))
        self.style.configure('Subtitle.TLabel', background='#0b1b2b', foreground='#7ec8ff', font=('Segoe UI', 10, 'bold'))
        self.style.configure('Section.TLabel', background='#ffffff', foreground='#143a57', font=('Segoe UI', 11, 'bold'))
        self.style.configure('Grid.TLabel', background='#ffffff', foreground='#1a2c3f', font=('Segoe UI', 10))
        self.style.configure('CardValue.TLabel', background='#ffffff', foreground='#0a2240', font=('Segoe UI', 17, 'bold'))
        self.style.configure('CardTitle.TLabel', background='#ffffff', foreground='#58738c', font=('Segoe UI', 9, 'bold'))
        self.style.configure('Widget.TEntry', fieldbackground='#f6fbff', foreground='#1a2c3f')
        self.style.configure('TCombobox', fieldbackground='#f6fbff', foreground='#1a2c3f')

        self.file_label = tk.StringVar(value='')
        self.model_label = tk.StringVar(value='')
        self.output_dir = tk.StringVar(value='')
        self.year_var = tk.StringVar(value='2026')
        self.month_var = tk.StringVar(value='08')
        self.date_selector = tk.StringVar(value='fecha_cancelacion')
        self.config = load_config(Path(__file__).resolve().parent / 'config.json')
        self.analysis_df = pd.DataFrame()
        self.last_result = {}

        self._build_ui()
        self._apply_config_to_fields()

    def _build_ui(self) -> None:
        top = tk.Frame(self, bg='#0b1b2b', height=92)
        top.pack(fill='x')
        tk.Label(top, text='Conversor de cobranzas ODE', bg='#0b1b2b', fg='#ffffff', font=('Segoe UI', 24, 'bold')).pack(anchor='w', padx=24, pady=(18, 4))
        tk.Label(top, text='Reporte de comprobantes cancelados → Formato de carga contable', bg='#0b1b2b', fg='#7ec8ff', font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=24)

        main = tk.Frame(self, bg='#edf4fb', padx=16, pady=14)
        main.pack(fill='both', expand=True)

        selector_frame = tk.LabelFrame(main, text='Archivos y periodo', bg='#ffffff', fg='#143a57', font=('Segoe UI', 11, 'bold'), padx=12, pady=10)
        selector_frame.pack(fill='x', pady=(0, 12))

        row1 = tk.Frame(selector_frame, bg='#ffffff')
        row1.pack(fill='x', pady=4)
        tk.Label(row1, text='Reporte de cancelados', bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 10), anchor='w', width=170)
        self.source_entry = tk.Entry(row1, width=70, textvariable=self.file_label, font=('Segoe UI', 10), relief='solid')
        self.source_entry.pack(side='left', fill='x', expand=True)
        tk.Button(row1, text='Seleccionar', command=self.select_source, bg='#1d4d82', fg='white', relief='flat', padx=10, pady=6).pack(side='left', padx=(8, 0))

        row2 = tk.Frame(selector_frame, bg='#ffffff')
        row2.pack(fill='x', pady=4)
        tk.Label(row2, text='Excel modelo (opcional)', bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 10), anchor='w', width=170)
        self.model_entry = tk.Entry(row2, width=70, textvariable=self.model_label, font=('Segoe UI', 10), relief='solid')
        self.model_entry.pack(side='left', fill='x', expand=True)
        tk.Button(row2, text='Seleccionar', command=self.select_model, bg='#1d4d82', fg='white', relief='flat', padx=10, pady=6).pack(side='left', padx=(8, 0))

        row3 = tk.Frame(selector_frame, bg='#ffffff')
        row3.pack(fill='x', pady=4)
        tk.Label(row3, text='Carpeta de salida', bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 10), anchor='w', width=170)
        self.output_entry = tk.Entry(row3, width=70, textvariable=self.output_dir, font=('Segoe UI', 10), relief='solid')
        self.output_entry.pack(side='left', fill='x', expand=True)
        tk.Button(row3, text='Seleccionar', command=self.select_output, bg='#1d4d82', fg='white', relief='flat', padx=10, pady=6).pack(side='left', padx=(8, 0))

        row4 = tk.Frame(selector_frame, bg='#ffffff')
        row4.pack(fill='x', pady=(8, 0))
        tk.Label(row4, text='Periodo contable', bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold')).pack(side='left', width=170, anchor='w')
        tk.Label(row4, text='Año', bg='#ffffff', fg='#0d2140').pack(side='left', padx=(0, 6))
        self.year_combo = ttk.Combobox(row4, textvariable=self.year_var, values=[str(y) for y in range(2024, 2036)], width=8, state='readonly')
        self.year_combo.pack(side='left', padx=(0, 16))
        tk.Label(row4, text='Mes', bg='#ffffff', fg='#0d2140').pack(side='left', padx=(0, 6))
        self.month_combo = ttk.Combobox(row4, textvariable=self.month_var, values=[f'{i:02d}' for i in range(1, 13)], width=6, state='readonly')
        self.month_combo.pack(side='left', padx=(0, 16))
        tk.Label(row4, text='Fecha registro', bg='#ffffff', fg='#0d2140').pack(side='left', padx=(0, 6))
        self.date_combo = ttk.Combobox(row4, textvariable=self.date_selector, values=['fecha_cancelacion', 'fecha_emi', 'ultimo_dia_mes'], width=20, state='readonly')
        self.date_combo.pack(side='left')

        config_frame = tk.LabelFrame(main, text='Configuración contable', bg='#ffffff', fg='#143a57', font=('Segoe UI', 11, 'bold'), padx=12, pady=10)
        config_frame.pack(fill='x', pady=(0, 12))

        fields = [
            ('Cuenta clientes', 'cuenta_clientes'),
            ('Sub diario', 'sub_diario'),
            ('Tipo anexo', 'tipo_anexo'),
            ('Tipo documento', 'tipo_documento'),
            ('Moneda', 'moneda'),
            ('Tipo conversión', 'tipo_conversion'),
            ('Doc. anulado', 'doc_anulado'),
            ('Glosa', 'glosa'),
        ]
        self.config_entries = {}
        for i, (label, key) in enumerate(fields):
            row = tk.Frame(config_frame, bg='#ffffff')
            row.pack(fill='x', pady=4)
            tk.Label(row, text=label, bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold'), width=18, anchor='w').pack(side='left')
            entry = tk.Entry(row, width=24, font=('Segoe UI', 10), relief='solid')
            entry.pack(side='left')
            self.config_entries[key] = entry

        bank_frame = tk.LabelFrame(main, text='Bancos', bg='#ffffff', fg='#143a57', font=('Segoe UI', 11, 'bold'), padx=12, pady=10)
        bank_frame.pack(fill='x', pady=(0, 12))
        self.bank_entries = {}
        for bank_name in ['INTERBANK', 'BBVA CONTINENTAL', 'DE LA NACION', 'OTROS']:
            row = tk.Frame(bank_frame, bg='#ffffff')
            row.pack(fill='x', pady=3)
            tk.Label(row, text=bank_name, bg='#ffffff', fg='#0d2140', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
            account = tk.Entry(row, width=16, font=('Segoe UI', 10), relief='solid')
            account.pack(side='left', padx=(0, 10))
            medio = tk.Entry(row, width=12, font=('Segoe UI', 10), relief='solid')
            medio.pack(side='left')
            self.bank_entries[bank_name] = {'cuenta': account, 'medio': medio}

        preview_frame = tk.LabelFrame(main, text='Vista previa', bg='#ffffff', fg='#143a57', font=('Segoe UI', 11, 'bold'), padx=10, pady=8)
        preview_frame.pack(fill='both', expand=True, pady=(0, 12))
        self.preview = ttk.Treeview(preview_frame, columns=('banco', 'serie', 'documento', 'importe', 'fecha'), show='headings', height=9)
        self.preview.heading('banco', text='Banco')
        self.preview.heading('serie', text='Serie')
        self.preview.heading('documento', text='Doc.')
        self.preview.heading('importe', text='Importe')
        self.preview.heading('fecha', text='Fecha')
        self.preview.column('banco', width=150)
        self.preview.column('serie', width=90)
        self.preview.column('documento', width=110)
        self.preview.column('importe', width=110, anchor='e')
        self.preview.column('fecha', width=120)
        scroll = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview.yview)
        self.preview.configure(yscrollcommand=scroll.set)
        self.preview.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        validation_frame = tk.LabelFrame(main, text='Validaciones y advertencias', bg='#ffffff', fg='#143a57', font=('Segoe UI', 11, 'bold'), padx=10, pady=8)
        validation_frame.pack(fill='x', pady=(0, 8))
        self.validation_box = tk.Text(validation_frame, height=8, width=100, wrap='word', bg='#fffdf5', fg='#1b2a39', font=('Segoe UI', 10))
        self.validation_box.pack(fill='both', expand=True)

        summary_frame = tk.Frame(main, bg='#edf4fb')
        summary_frame.pack(fill='x', pady=(0, 12))
        card_defs = [
            ('Registros encontrados', 'found'),
            ('Registros excluidos', 'excluded'),
            ('Registros procesados', 'processed'),
            ('Total cobrado', 'total'),
            ('Diferencia debe/haber', 'balance'),
            ('Estado final', 'status'),
        ]
        self.cards = {}
        for idx, (label, key) in enumerate(card_defs):
            card = tk.Frame(summary_frame, bg='#ffffff', highlightbackground='#dfeaf5', highlightthickness=1, padx=12, pady=12)
            card.grid(row=0, column=idx, padx=8, pady=4, sticky='nsew')
            tk.Label(card, text=label, bg='#ffffff', fg='#56708a', font=('Segoe UI', 9, 'bold')).pack(anchor='w')
            value = tk.Label(card, text='0', bg='#ffffff', fg='#102a4a', font=('Segoe UI', 18, 'bold'))
            value.pack(anchor='w', pady=(8, 0))
            self.cards[key] = value

        actions = tk.Frame(main, bg='#edf4fb')
        actions.pack(fill='x', pady=(8, 10))
        self.progress = ttk.Progressbar(actions, orient='horizontal', mode='determinate', length=420)
        self.progress.pack(side='left', padx=(0, 12))
        self.status_label = tk.Label(actions, text='Listo', bg='#edf4fb', fg='#204a6d', font=('Segoe UI', 10, 'bold'))
        self.status_label.pack(side='left', padx=(4, 10))
        tk.Button(actions, text='Analizar archivos', command=self.analyze_files, bg='#1d4d82', fg='white', font=('Segoe UI', 10, 'bold'), padx=16, pady=8, relief='flat').pack(side='left', padx=5)
        tk.Button(actions, text='Generar carga', command=self.generate_load, bg='#2a9d8f', fg='white', font=('Segoe UI', 10, 'bold'), padx=16, pady=8, relief='flat').pack(side='left', padx=5)
        tk.Button(actions, text='Abrir carpeta de salida', command=self.open_output_folder, bg='#3a6ea5', fg='white', font=('Segoe UI', 10, 'bold'), padx=16, pady=8, relief='flat').pack(side='left', padx=5)
        tk.Button(actions, text='Restablecer', command=self.reset_all, bg='#cbcdd2', fg='#102a4a', font=('Segoe UI', 10, 'bold'), padx=16, pady=8, relief='flat').pack(side='left', padx=5)
        tk.Button(actions, text='Guardar configuración', command=self.save_config_fields, bg='#e0b03d', fg='#102a4a', font=('Segoe UI', 10, 'bold'), padx=16, pady=8, relief='flat').pack(side='left', padx=5)

        self.output_dir.set(str(Path(__file__).resolve().parent / 'salida'))
        self.source_entry.focus_set()

    def _apply_config_to_fields(self) -> None:
        for key, entry in self.config_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, str(self.config.get(key, DEFAULT_CONFIG.get(key, ''))))

        for bank_name, fields in self.bank_entries.items():
            bank_cfg = self.config.get('bancos', {}).get(bank_name, {'cuenta_contable': '10100011', 'medio_pago': '008'})
            fields['cuenta'].delete(0, tk.END)
            fields['cuenta'].insert(0, bank_cfg.get('cuenta_contable', '10100011'))
            fields['medio'].delete(0, tk.END)
            fields['medio'].insert(0, bank_cfg.get('medio_pago', '008'))

    def select_source(self) -> None:
        path = filedialog.askopenfilename(title='Seleccionar reporte de cancelados', filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')])
        if path:
            self.file_label.set(path)

    def select_model(self) -> None:
        path = filedialog.askopenfilename(title='Seleccionar modelo contable', filetypes=[('Excel', '*.xls *.xlsx'), ('Todos', '*.*')])
        if path:
            self.model_label.set(path)

    def select_output(self) -> None:
        folder = filedialog.askdirectory(title='Seleccionar carpeta de salida')
        if folder:
            self.output_dir.set(folder)

    def save_config_fields(self) -> None:
        cfg = self.config.copy()
        for key, entry in self.config_entries.items():
            cfg[key] = entry.get().strip()
        cfg['bancos'] = cfg.get('bancos', {})
        for bank_name, fields in self.bank_entries.items():
            cfg['bancos'][bank_name] = {
                'cuenta_contable': fields['cuenta'].get().strip() or '10100011',
                'medio_pago': fields['medio'].get().strip() or '008',
            }
        self.config = cfg
        save_config(Path(__file__).resolve().parent / 'config.json', cfg)
        messagebox.showinfo('Configuración guardada', 'La configuración contable se guardó correctamente.')

    def analyze_files(self) -> None:
        source_path = self.file_label.get()
        if not source_path:
            messagebox.showwarning('Archivo requerido', 'Debe seleccionar un reporte de cancelados.')
            return
        self.progress['value'] = 10
        self.status_label.config(text='Analizando reporte...')
        thread = threading.Thread(target=self._analyze_worker, args=(source_path,), daemon=True)
        thread.start()

    def _analyze_worker(self, source_path: str) -> None:
        try:
            df = load_ode_report(source_path)
            excluded = df[df['serie'].astype(str).str.startswith('0001')].copy()
            processed = df[~df['serie'].astype(str).str.startswith('0001')].copy()
            errors, warnings, info = validate_report(processed)
            self.analysis_df = processed
            self.last_result = {'processed': processed, 'excluded': excluded, 'errors': errors, 'warnings': warnings, 'info': info}
            self.after(0, lambda: self._refresh_analysis(processed, excluded, errors, warnings, info))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror('Error', str(exc)))
            self.after(0, lambda: self.status_label.config(text='Error al analizar'))
        finally:
            self.after(0, lambda: self.progress.configure(value=100))

    def _refresh_analysis(self, processed: pd.DataFrame, excluded: pd.DataFrame, errors: list, warnings: list, info: list) -> None:
        self.status_label.config(text='Analisis finalizado')
        preview = processed[['banco', 'serie', 'documento', 'importe_cancelado', 'fecha_cancelacion']].head(100).copy()
        preview.columns = ['banco', 'serie', 'documento', 'importe', 'fecha']
        preview['importe'] = preview['importe'].apply(lambda x: f'{float(x):,.2f}' if pd.notna(x) else '0.00')
        self.preview.delete(*self.preview.get_children())
        for _, row in preview.iterrows():
            self.preview.insert('', 'end', values=(row['banco'], row['serie'], row['documento'], row['importe'], str(row['fecha'])) )

        total = float(processed['importe_cancelado'].sum()) if 'importe_cancelado' in processed.columns else 0
        self.cards['found'].config(text=str(len(processed) + len(excluded)))
        self.cards['excluded'].config(text=str(len(excluded)))
        self.cards['processed'].config(text=str(len(processed)))
        self.cards['total'].config(text=f'S/{total:,.2f}')
        self.cards['balance'].config(text='0.00')
        self.cards['status'].config(text='LISTO' if not errors else 'REVISAR')

        validation_text = '\n'.join([f'[ERROR] {m}' for m in errors] + [f'[ADVERTENCIA] {m}' for m in warnings] + [f'[INFO] {m}' for m in info])
        self.validation_box.delete('1.0', tk.END)
        self.validation_box.insert('1.0', validation_text)

    def generate_load(self) -> None:
        source_path = self.file_label.get()
        output_dir = self.output_dir.get() or str(Path(__file__).resolve().parent / 'salida')
        if not source_path:
            messagebox.showwarning('Archivo requerido', 'Debe seleccionar un reporte antes de generar la carga.')
            return
        self.progress['value'] = 5
        self.status_label.config(text='Generando carga...')
        cfg = self.config.copy()
        for key, entry in self.config_entries.items():
            cfg[key] = entry.get().strip()
        for bank_name, fields in self.bank_entries.items():
            cfg.setdefault('bancos', {})[bank_name] = {
                'cuenta_contable': fields['cuenta'].get().strip() or '10100011',
                'medio_pago': fields['medio'].get().strip() or '008',
            }
        self.config = cfg
        thread = threading.Thread(target=self._generate_worker, args=(source_path, output_dir, cfg), daemon=True)
        thread.start()

    def _generate_worker(self, source_path: str, output_dir: str, cfg: dict) -> None:
        try:
            result = process_report(source_path, output_dir, cfg, self.date_selector.get())
            self.after(0, lambda: self._finish_generation(result))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror('Error al generar', str(exc)))
            self.after(0, lambda: self.status_label.config(text='Error de generación'))
        finally:
            self.after(0, lambda: self.progress.configure(value=100))

    def _finish_generation(self, result: dict) -> None:
        self.status_label.config(text='Carga generada')
        self.last_result = result
        summary = result['summary']
        self.cards['found'].config(text=str(summary.get('cantidad_original', 0)))
        self.cards['excluded'].config(text=str(summary.get('cantidad_excluida', 0)))
        self.cards['processed'].config(text=str(summary.get('cantidad_procesada', 0)))
        self.cards['total'].config(text=f"S/{float(summary.get('total_procesado', 0.0)):.2f}")
        self.cards['balance'].config(text=f"S/{float(summary.get('diferencia', 0.0)):.2f}")
        self.cards['status'].config(text=str(summary.get('resultado', 'LISTO')))
        self.validation_box.delete('1.0', tk.END)
        text = '\n'.join([f'[ADVERTENCIA] {m}' for m in result.get('warnings', [])] + [f'[INFO] {m}' for m in result.get('info', [])])
        self.validation_box.insert('1.0', text)
        messagebox.showinfo('Carga generada', f'Archivo generado: {result["output_file"]}\nReporte: {result["report_file"]}')

    def open_output_folder(self) -> None:
        folder = self.output_dir.get() or str(Path(__file__).resolve().parent / 'salida')
        Path(folder).mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(folder)
        except Exception:
            try:
                webbrowser.open(folder)
            except Exception:
                pass

    def reset_all(self) -> None:
        self.file_label.set('')
        self.model_label.set('')
        self.output_dir.set(str(Path(__file__).resolve().parent / 'salida'))
        self.year_var.set('2026')
        self.month_var.set('08')
        self.date_selector.set('fecha_cancelacion')
        self.validation_box.delete('1.0', tk.END)
        self.preview.delete(*self.preview.get_children())
        self.progress['value'] = 0
        self.status_label.config(text='Listo')
        self.config = load_config(Path(__file__).resolve().parent / 'config.json')
        self._apply_config_to_fields()


if __name__ == '__main__':
    app = ConversorODEApp()
    app.mainloop()
