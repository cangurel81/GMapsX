import os
import sys
import threading
import subprocess
from datetime import datetime
from pathlib import Path


def _resource_path(relative_path):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scraper import GoogleMapsScraper

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_t = ctk.ThemeManager.theme
_t["CTkCheckBox"]["checkmark_color"] = ["#DCE4EE", "#8b949e"]
_t["DropdownMenu"]["fg_color"]       = ["gray90", "#161b22"]
_t["DropdownMenu"]["hover_color"]    = ["gray75", "#21262d"]

BG_PANEL   = "#161b22"
BG_SURFACE = "#21262d"
BG_INPUT   = "#2d333b"
BG_HOVER   = "#30363d"
FG_PRIMARY = "#e6edf3"
FG_MUTED   = "#8b949e"

FONT_MAIN  = ("Segoe UI", 12)
FONT_BOLD  = ("Segoe UI", 12, "bold")
FONT_MONO  = ("Consolas", 11)


def entry(parent, textvariable) -> ctk.CTkEntry:
    return ctk.CTkEntry(
        parent, textvariable=textvariable,
        fg_color=BG_INPUT, text_color=FG_PRIMARY,
        font=FONT_MAIN, corner_radius=4,
        border_width=1, border_color=BG_SURFACE,
    )


C_HEADER      = "21262D"
C_HEADER_DARK = "161B22"
C_BG_EVEN     = "F0F2F5"
C_BG_ODD      = "FFFFFF"
C_TEXT_DARK   = "1F2937"
C_TEXT_GRAY   = "6B7280"
C_LINK        = "2563EB"
C_GREEN       = "059669"
C_BORDER      = "E5E7EB"
C_BORDER_HDR  = "161B22"

COLUMNS = [
    "İsim", "Kategori", "Puan", "Yorum Sayısı", "Adres",
    "Telefon", "Website", "Google Maps URL",
]

COL_WIDTHS = {
    "İsim": 32, "Kategori": 18, "Puan": 8, "Yorum Sayısı": 13,
    "Adres": 44, "Telefon": 18, "Website": 38, "Google Maps URL": 22,
}


def _border(color=C_BORDER, style="thin") -> Border:
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


class ExcelExporter:

    def export(self, data: list[dict], filepath: str,
               query: str = "", location: str = "") -> Path:
        wb = openpyxl.Workbook()
        self._build_results_sheet(wb.active, data, query, location)
        self._build_summary_sheet(wb, data, query, location)
        path = Path(filepath)
        wb.save(path)
        return path

    def _build_results_sheet(self, ws, data, query, location):
        ws.title = "Sonuçlar"
        ws.sheet_view.showGridLines = False

        n_cols = len(COLUMNS)

        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=n_cols)
        title_cell = ws.cell(row=1, column=1,
                             value="GMapsX – Sonuçlar")
        title_cell.font = Font(name="Segoe UI", bold=True, size=15,
                               color="FFFFFF")
        title_cell.fill = _fill(C_HEADER)
        title_cell.alignment = Alignment(horizontal="left", vertical="center",
                                         indent=2)
        ws.row_dimensions[1].height = 38

        ws.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=n_cols // 2)
        meta = ws.cell(
            row=2, column=1,
            value=f"Arama: {query}   |   Konum: {location}   |   "
                  f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        meta.font = Font(name="Segoe UI", size=9, color=C_TEXT_GRAY)
        meta.fill = _fill(C_BG_EVEN)
        meta.alignment = Alignment(horizontal="left", vertical="center",
                                   indent=2)

        ws.merge_cells(start_row=2, start_column=n_cols // 2 + 1,
                       end_row=2, end_column=n_cols)
        count_cell = ws.cell(row=2, column=n_cols // 2 + 1,
                             value=f"Toplam  {len(data)}  kayıt")
        count_cell.font = Font(name="Segoe UI", bold=True, size=10,
                               color=C_HEADER_DARK)
        count_cell.fill = _fill(C_BG_EVEN)
        count_cell.alignment = Alignment(horizontal="right", vertical="center",
                                         indent=2)
        ws.row_dimensions[2].height = 22

        ws.row_dimensions[3].height = 6

        HDR_ROW = 4
        thin_hdr = Side(style="medium", color=C_HEADER_DARK)
        for col_idx, col_name in enumerate(COLUMNS, 1):
            cell = ws.cell(row=HDR_ROW, column=col_idx, value=col_name)
            cell.font = Font(name="Segoe UI", bold=True, size=9.5,
                             color="FFFFFF")
            cell.fill = _fill(C_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=False)
            cell.border = Border(
                top=thin_hdr, bottom=thin_hdr,
                left=Side(style="thin", color=C_HEADER_DARK),
                right=Side(style="thin", color=C_HEADER_DARK),
            )
        ws.row_dimensions[HDR_ROW].height = 28

        for r_idx, record in enumerate(data, HDR_ROW + 1):
            even = (r_idx - HDR_ROW) % 2 == 0
            bg = C_BG_EVEN if even else C_BG_ODD
            b = _border(C_BORDER)

            for c_idx, col_name in enumerate(COLUMNS, 1):
                value = record.get(col_name, "")
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = Font(name="Segoe UI", size=9, color=C_TEXT_DARK)
                cell.fill = _fill(bg)
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                cell.border = b

                if col_name == "Puan" and value:
                    cell.alignment = Alignment(horizontal="center",
                                               vertical="center")

                if col_name == "Yorum Sayısı" and value:
                    cell.alignment = Alignment(horizontal="right",
                                               vertical="center")

                if col_name == "Website" and value and value.startswith("http"):
                    cell.hyperlink = value
                    cell.font = Font(name="Segoe UI", size=9,
                                     color=C_LINK, underline="single")

                if col_name == "Google Maps URL" and value and value.startswith("http"):
                    cell.hyperlink = value
                    cell.value = "Haritada Gör"
                    cell.font = Font(name="Segoe UI", size=9,
                                     color=C_GREEN, underline="single")
                    cell.alignment = Alignment(horizontal="center",
                                               vertical="center")

            ws.row_dimensions[r_idx].height = 21

        for c_idx, col_name in enumerate(COLUMNS, 1):
            ws.column_dimensions[get_column_letter(c_idx)].width = (
                COL_WIDTHS.get(col_name, 15)
            )

        ws.freeze_panes = ws.cell(row=HDR_ROW + 1, column=1)

        last_col_letter = get_column_letter(n_cols)
        ws.auto_filter.ref = (
            f"A{HDR_ROW}:{last_col_letter}{HDR_ROW + len(data)}"
        )

    def _build_summary_sheet(self, wb, data, query, location):
        ws = wb.create_sheet("Özet")
        ws.sheet_view.showGridLines = False

        ws.merge_cells("A1:C1")
        t = ws.cell(row=1, column=1, value="Arama Özeti")
        t.font = Font(name="Segoe UI", bold=True, size=14, color="FFFFFF")
        t.fill = _fill(C_HEADER)
        t.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        ws.row_dimensions[1].height = 36

        rows = [
            ("Arama Terimi", query),
            ("Konum", location),
            ("Tarih", datetime.now().strftime("%d.%m.%Y %H:%M")),
            ("", ""),
            ("Toplam Kayıt", len(data)),
            ("Telefonu Olanlar", sum(1 for d in data if d.get("Telefon"))),
            ("Websitesi Olanlar", sum(1 for d in data if d.get("Website"))),
            ("Puanı Olanlar", sum(1 for d in data if d.get("Puan"))),
            ("Adresi Olanlar", sum(1 for d in data if d.get("Adres"))),
        ]

        cats: dict[str, int] = {}
        for d in data:
            cat = d.get("Kategori", "—") or "—"
            cats[cat] = cats.get(cat, 0) + 1

        for r_off, (label, value) in enumerate(rows, 2):
            if not label:
                ws.row_dimensions[r_off].height = 8
                continue

            lbl = ws.cell(row=r_off, column=1, value=label)
            lbl.font = Font(name="Segoe UI", size=10, color=C_TEXT_GRAY)
            lbl.alignment = Alignment(horizontal="left", vertical="center",
                                      indent=1)
            lbl.fill = _fill(C_BG_EVEN if r_off % 2 == 0 else "FFFFFF")

            val = ws.cell(row=r_off, column=2, value=value)
            val.font = Font(name="Segoe UI", bold=True, size=11,
                            color=C_HEADER_DARK)
            val.alignment = Alignment(horizontal="left", vertical="center",
                                      indent=1)
            val.fill = _fill(C_BG_EVEN if r_off % 2 == 0 else "FFFFFF")
            ws.row_dimensions[r_off].height = 22

        cat_start = len(rows) + 4
        ws.cell(row=cat_start, column=1,
                value="Kategori Dağılımı").font = Font(
            name="Segoe UI", bold=True, size=10, color="FFFFFF")
        ws.cell(row=cat_start, column=1).fill = _fill(C_HEADER)
        ws.cell(row=cat_start, column=2, value="Adet").font = Font(
            name="Segoe UI", bold=True, size=10, color="FFFFFF")
        ws.cell(row=cat_start, column=2).fill = _fill(C_HEADER)

        for i, (cat, cnt) in enumerate(
            sorted(cats.items(), key=lambda x: -x[1]), cat_start + 1
        ):
            bg = C_BG_EVEN if i % 2 == 0 else "FFFFFF"
            c1 = ws.cell(row=i, column=1, value=cat)
            c1.font = Font(name="Segoe UI", size=9)
            c1.fill = _fill(bg)
            c2 = ws.cell(row=i, column=2, value=cnt)
            c2.font = Font(name="Segoe UI", size=9, bold=True,
                           color=C_HEADER_DARK)
            c2.fill = _fill(bg)
            c2.alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 16
        ws.column_dimensions["C"].width = 12


class App(ctk.CTk):

    LANGUAGES = {
        "Türkçe (tr)": "tr",
        "English (en)": "en",
        "Deutsch (de)": "de",
        "Français (fr)": "fr",
        "Español (es)": "es",
        "Italiano (it)": "it",
        "Português (pt)": "pt",
        "Nederlands (nl)": "nl",
        "Русский (ru)": "ru",
        "العربية (ar)": "ar",
        "中文 (zh)": "zh",
        "日本語 (ja)": "ja",
        "한국어 (ko)": "ko",
        "हिन्दी (hi)": "hi",
        "Bahasa Indonesia (id)": "id",
        "ภาษาไทย (th)": "th",
        "Tiếng Việt (vi)": "vi",
        "Polski (pl)": "pl",
        "Svenska (sv)": "sv",
        "Norsk (no)": "no",
        "Dansk (da)": "da",
        "Suomi (fi)": "fi",
        "Čeština (cs)": "cs",
        "Română (ro)": "ro",
        "Magyar (hu)": "hu",
        "Ελληνικά (el)": "el",
        "עברית (he)": "he",
    }

    def __init__(self):
        super().__init__()
        self.title("GMapsX")
        self.geometry("600x600")
        self.resizable(False, False)
        try:
            self.iconbitmap(_resource_path("maps.ico"))
            self.after(200, lambda: self.iconbitmap(_resource_path("maps.ico")))
        except Exception:
            pass

        self._data: list[dict] = []
        self._scraper: GoogleMapsScraper | None = None
        self._thread: threading.Thread | None = None
        self._running = False

        self._build_ui()

    def _build_ui(self):
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=10, pady=(22, 8))

        left = ctk.CTkFrame(content, fg_color=BG_PANEL, width=280,
                            corner_radius=6)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self._build_form(left)

        right = ctk.CTkFrame(content, fg_color=BG_PANEL, width=260,
                             corner_radius=6)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        self._build_buttons(right)

        self._build_statusbar()

    def _build_form(self, pnl):
        section = ctk.CTkFrame(pnl, fg_color="transparent")
        section.pack(fill="both", expand=True, padx=14, pady=(16, 0))

        sep = tk.Frame(section, bg="#1c2128", height=1)
        sep.pack(fill="x")

        ctk.CTkLabel(section, text="ARAMA VE AYARLAR",
                      text_color=FG_PRIMARY,
                      font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(4, 10))

        ctk.CTkLabel(section, text="İşletme türü / anahtar kelime",
                      text_color=FG_MUTED, font=FONT_MAIN).pack(anchor="w")
        self._query_var = tk.StringVar(value="restoran")
        entry(section, self._query_var).pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(section, text="Konum",
                      text_color=FG_MUTED, font=FONT_MAIN).pack(anchor="w")
        self._loc_var = tk.StringVar(value="İstanbul")
        entry(section, self._loc_var).pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(section, text="Maksimum sonuç",
                      text_color=FG_MUTED, font=FONT_MAIN).pack(anchor="w")
        row = ctk.CTkFrame(section, fg_color="transparent")
        row.pack(fill="x", pady=(2, 6))
        self._max_var = tk.IntVar(value=20)
        self._max_lbl = ctk.CTkLabel(row, text="20", text_color=FG_PRIMARY,
                                      font=FONT_BOLD, width=30)
        self._max_lbl.pack(side="right")
        ctk.CTkSlider(
            row, from_=5, to=200, number_of_steps=195,
            variable=self._max_var, orientation="horizontal",
            fg_color="#1c2128", progress_color=FG_MUTED,
            button_color=BG_HOVER, button_hover_color=BG_SURFACE,
            command=lambda v: self._max_lbl.configure(text=str(int(round(v))))
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkLabel(section, text="Eş zamanlı thread",
                      text_color=FG_MUTED, font=FONT_MAIN).pack(anchor="w")
        row2 = ctk.CTkFrame(section, fg_color="transparent")
        row2.pack(fill="x", pady=(2, 6))
        self._thread_var = tk.IntVar(value=3)
        self._thread_lbl = ctk.CTkLabel(row2, text="3", text_color=FG_PRIMARY,
                                         font=FONT_BOLD, width=30)
        self._thread_lbl.pack(side="right")
        ctk.CTkSlider(
            row2, from_=1, to=20, number_of_steps=19,
            variable=self._thread_var, orientation="horizontal",
            fg_color="#1c2128", progress_color=FG_MUTED,
            button_color=BG_HOVER, button_hover_color=BG_SURFACE,
            command=lambda v: self._thread_lbl.configure(text=str(int(round(v))))
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkLabel(section, text="Arama dili",
                      text_color=FG_MUTED, font=FONT_MAIN).pack(anchor="w")
        self._lang_var = tk.StringVar(value="Türkçe (tr)")
        ctk.CTkComboBox(
            section, variable=self._lang_var, values=list(self.LANGUAGES.keys()),
            fg_color=BG_INPUT, button_color=BG_SURFACE,
            button_hover_color=BG_HOVER, dropdown_fg_color=BG_INPUT,
            dropdown_hover_color=BG_HOVER, dropdown_text_color=FG_PRIMARY,
            text_color=FG_PRIMARY, corner_radius=4,
        ).pack(fill="x", pady=(2, 8))

        self._headless_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            section, text="Arka planda çalıştır (headless)",
            variable=self._headless_var,
            fg_color=BG_INPUT, hover_color=BG_HOVER,
            text_color=FG_MUTED, font=FONT_MAIN,
        ).pack(anchor="w", pady=(2, 0))


    def _build_buttons(self, parent):
        section = ctk.CTkFrame(parent, fg_color="transparent")
        section.pack(fill="both", expand=True, padx=14, pady=(16, 0))

        tk.Frame(section, bg="#1c2128", height=1).pack(fill="x")

        ctk.CTkLabel(section, text="İŞLEMLER",
                      text_color=FG_PRIMARY,
                      font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(4, 14))

        btn_pad = {"fill": "x", "pady": 3}

        self._start_btn = ctk.CTkButton(
            section, text="▶   Aramayı Başlat", command=self.start_scraping,
            fg_color=BG_INPUT, hover_color=BG_HOVER,
            text_color=FG_PRIMARY, font=FONT_BOLD, corner_radius=4,
        )
        self._start_btn.pack(**btn_pad)

        self._stop_btn = ctk.CTkButton(
            section, text="⏹   Durdur", command=self.stop_scraping,
            fg_color=BG_INPUT, hover_color=BG_HOVER,
            text_color=FG_PRIMARY, font=FONT_BOLD, corner_radius=4,
            state="disabled",
        )
        self._stop_btn.pack(**btn_pad)

        self._export_btn = ctk.CTkButton(
            section, text="📊   Excel'e Aktar", command=self.export_excel,
            fg_color=BG_SURFACE, hover_color=BG_HOVER,
            text_color=FG_PRIMARY, font=FONT_BOLD, corner_radius=4,
            state="disabled",
        )
        self._export_btn.pack(**btn_pad)

        self._export_pdf_btn = ctk.CTkButton(
            section, text="📄   PDF'e Aktar", command=self.export_pdf,
            fg_color=BG_SURFACE, hover_color=BG_HOVER,
            text_color=FG_PRIMARY, font=FONT_BOLD, corner_radius=4,
            state="disabled",
        )
        self._export_pdf_btn.pack(**btn_pad)

        self._export_txt_btn = ctk.CTkButton(
            section, text="📝   TXT'e Aktar", command=self.export_txt,
            fg_color=BG_SURFACE, hover_color=BG_HOVER,
            text_color=FG_PRIMARY, font=FONT_BOLD, corner_radius=4,
            state="disabled",
        )
        self._export_txt_btn.pack(**btn_pad)

        ctk.CTkButton(
            section, text="🗑   Sonuçları Temizle", command=self.clear_results,
            fg_color=BG_SURFACE, hover_color=BG_HOVER,
            text_color=FG_MUTED, font=FONT_MAIN, corner_radius=4,
        ).pack(**btn_pad)

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=BG_PANEL, height=50,
                           corner_radius=4)
        bar.pack(fill="x", padx=10, pady=(0, 10))
        bar.pack_propagate(False)

        self._progress = ctk.CTkProgressBar(
            bar, mode="indeterminate",
            fg_color=BG_SURFACE, progress_color=FG_MUTED,
            corner_radius=2,
        )
        self._progress.pack(fill="x", padx=8, pady=(6, 2))

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="x", padx=10)

        self._status_var = tk.StringVar(
            value="Hazır."
        )
        ctk.CTkLabel(inner, textvariable=self._status_var,
                      text_color=FG_MUTED, font=FONT_MONO,
                      anchor="w").pack(side="left")

        ctk.CTkLabel(inner, text="WebAdHere Software",
                      text_color=FG_MUTED, font=FONT_MONO,
                      anchor="e").pack(side="right")

    def _set_status(self, msg: str):
        self.after(0, lambda: self._status_var.set(msg))

    def start_scraping(self):
        if self._running:
            return

        query = self._query_var.get().strip()
        loc   = self._loc_var.get().strip()

        if not query:
            messagebox.showwarning("Eksik bilgi", "Lütfen arama terimi girin.")
            return
        if not loc:
            messagebox.showwarning("Eksik bilgi", "Lütfen konum girin.")
            return

        self.clear_results()
        self._running = True

        self._start_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._export_btn.configure(state="disabled")
        self._export_pdf_btn.configure(state="disabled")
        self._export_txt_btn.configure(state="disabled")
        self._progress.start()

        max_res   = int(self._max_var.get())
        threads   = int(self._thread_var.get())
        headless  = self._headless_var.get()
        lang_key  = self._lang_var.get()
        lang_code = self.LANGUAGES.get(lang_key, "tr")

        def _run():
            try:
                self._scraper = GoogleMapsScraper(
                    headless=headless,
                    on_result=lambda d, n: self._data.append(d),
                    on_log=self._set_status,
                )
                self._set_status("Chrome başlatılıyor…")
                self._scraper.setup_driver()
                self._scraper.search(
                    query=query,
                    location=loc,
                    max_results=max_res,
                    language=lang_code,
                    threads=threads,
                )
                self.after(0, self._on_done)
            except Exception as exc:
                self.after(0, lambda e=exc: self._on_error(e))

        self._thread = threading.Thread(target=_run, daemon=True)
        self._thread.start()

    def stop_scraping(self):
        if self._scraper:
            self._scraper.running = False
        self._running = False
        self._set_status("Durduruluyor…")
        self.after(1500, self._on_done)

    def _on_done(self):
        self._running = False
        self._progress.stop()
        self._start_btn.configure(state="normal")
        self._stop_btn.configure(state="disabled")
        n = len(self._data)
        if n:
            self._export_btn.configure(state="normal")
            self._export_pdf_btn.configure(state="normal")
            self._export_txt_btn.configure(state="normal")
        self._set_status(f"Tamamlandı ✓  {n} kayıt bulundu.")
        if self._scraper:
            threading.Thread(target=self._scraper.close, daemon=True).start()
            self._scraper = None

    def _on_error(self, exc: Exception):
        self._on_done()
        err = str(exc)
        self._set_status(f"Hata: {err[:120]}")
        messagebox.showerror("Hata", f"Scraping sırasında sorun oluştu:\n\n{err}")

    def clear_results(self):
        self._data.clear()

    def export_excel(self):
        if not self._data:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        q   = self._query_var.get().strip().replace(" ", "_")
        loc = self._loc_var.get().strip().replace(" ", "_")
        ts  = datetime.now().strftime("%Y%m%d_%H%M")
        default_name = f"{q}_{loc}_{ts}.xlsx"

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyası", "*.xlsx"),
                       ("Tüm Dosyalar", "*.*")],
            initialfile=default_name,
            title="Excel Dosyasını Kaydet",
        )
        if not path:
            return

        try:
            exporter = ExcelExporter()
            saved = exporter.export(
                self._data, path,
                query=self._query_var.get(),
                location=self._loc_var.get(),
            )
            messagebox.showinfo(
                "Başarılı",
                f"Veriler kaydedildi:\n{saved}\n\nToplam {len(self._data)} kayıt."
            )
            try:
                subprocess.Popen(
                    ["explorer", "/select,", str(saved).replace("/", "\\")]
                )
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror(
                "Excel Hatası",
                f"Dosya oluşturulurken hata:\n{exc}"
            )

    def export_txt(self):
        if not self._data:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        q   = self._query_var.get().strip().replace(" ", "_")
        loc = self._loc_var.get().strip().replace(" ", "_")
        ts  = datetime.now().strftime("%Y%m%d_%H%M")
        default_name = f"{q}_{loc}_{ts}.txt"

        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Metin Dosyası", "*.txt"),
                       ("Tüm Dosyalar", "*.*")],
            initialfile=default_name,
            title="TXT Dosyasını Kaydet",
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\t".join(COLUMNS) + "\n")
                for rec in self._data:
                    row = [str(rec.get(c, "")) for c in COLUMNS]
                    f.write("\t".join(row) + "\n")
            messagebox.showinfo(
                "Başarılı",
                f"Veriler kaydedildi:\n{path}\n\nToplam {len(self._data)} kayıt."
            )
            try:
                subprocess.Popen(
                    ["explorer", "/select,", str(path).replace("/", "\\")]
                )
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror(
                "TXT Hatası",
                f"Dosya oluşturulurken hata:\n{exc}"
            )

    def export_pdf(self):
        if not self._data:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        try:
            from fpdf import FPDF
        except ImportError:
            messagebox.showerror(
                "Eksik Kütüphane",
                "PDF export için 'fpdf2' kütüphanesi gerekli.\n\nKurulum: pip install fpdf2"
            )
            return

        q   = self._query_var.get().strip().replace(" ", "_")
        loc = self._loc_var.get().strip().replace(" ", "_")
        ts  = datetime.now().strftime("%Y%m%d_%H%M")
        default_name = f"{q}_{loc}_{ts}.pdf"

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf"), ("Tüm Dosyalar", "*.*")],
            initialfile=default_name,
            title="PDF Dosyasını Kaydet",
        )
        if not path:
            return

        try:
            pdf = FPDF()
            pdf.add_page()

            font_paths = [
                "C:\\Windows\\Fonts\\segoeui.ttf",
                "C:\\Windows\\Fonts\\calibri.ttf",
                "C:\\Windows\\Fonts\\arial.ttf",
            ]
            font_name = "Custom"
            name_map = {"segoeui": "segoeuib", "calibri": "calibrib", "arial": "arialbd"}
            for fp in font_paths:
                if os.path.exists(fp):
                    pdf.add_font(font_name, "", fp, uni=True)
                    base = os.path.splitext(os.path.basename(fp))[0]
                    bold_name = name_map.get(base)
                    if bold_name:
                        bold_fp = os.path.join(os.path.dirname(fp), bold_name + ".ttf")
                        if os.path.exists(bold_fp):
                            pdf.add_font(font_name, "B", bold_fp, uni=True)
                        else:
                            pdf.add_font(font_name, "B", fp, uni=True)
                    break
            else:
                font_name = "Helvetica"

            pdf.set_font(font_name, "B", 14)
            pdf.cell(0, 10, "GMapsX - Sonuclar", ln=True)

            pdf.set_font(font_name, "", 8)
            meta = f"Arama: {self._query_var.get()}  |  Konum: {self._loc_var.get()}  |  Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            pdf.cell(0, 6, meta, ln=True)
            pdf.cell(0, 6, f"Toplam {len(self._data)} kayit", ln=True)
            pdf.ln(4)

            col_w = [50, 30, 12, 18, 55, 30, 50, 35]
            pdf.set_font(font_name, "B", 7)
            for i, col_name in enumerate(COLUMNS):
                pdf.cell(col_w[i], 7, col_name, border=1)
            pdf.ln()

            pdf.set_font(font_name, "", 7)
            for rec in self._data:
                for i, col_name in enumerate(COLUMNS):
                    val = str(rec.get(col_name, ""))[:40]
                    pdf.cell(col_w[i], 6, val, border=1)
                pdf.ln()
                if pdf.get_y() > 270:
                    pdf.add_page()
                    pdf.set_font(font_name, "B", 7)
                    for i, col_name in enumerate(COLUMNS):
                        pdf.cell(col_w[i], 7, col_name, border=1)
                    pdf.ln()
                    pdf.set_font(font_name, "", 7)

            pdf.output(path)
            messagebox.showinfo(
                "Başarılı",
                f"Veriler kaydedildi:\n{path}\n\nToplam {len(self._data)} kayıt."
            )
            try:
                subprocess.Popen(
                    ["explorer", "/select,", str(path).replace("/", "\\")]
                )
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror(
                "PDF Hatası",
                f"Dosya oluşturulurken hata:\n{exc}"
            )


if __name__ == "__main__":
    app = App()
    app.mainloop()
