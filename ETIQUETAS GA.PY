#!/usr/bin/env python3
"""
Generador de Etiquetas de Ubicación v2.1
Tamaño: 51.20 × 25.00 mm (EVA58)
Requiere: PyQt6, reportlab, openpyxl
"""

import sys
import os
import json
import shutil
import tempfile
import urllib.request
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QTextEdit, QPushButton, QSpinBox,
    QRadioButton, QButtonGroup, QComboBox, QCheckBox, QFileDialog,
    QMessageBox, QFrame, QGroupBox, QSizePolicy, QScrollArea,
    QStatusBar, QSplitter, QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressDialog,
)
from PyQt6.QtCore import Qt, QRectF, QSize, QObject, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import (
    QFont, QColor, QPainter, QPen, QBrush, QFontMetrics, QPalette, QIcon
)

from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code128

import openpyxl

# ── Constantes ────────────────────────────────────────────────────────────────
LABEL_W_MM = 51.20
LABEL_H_MM = 25.00
LABEL_W = LABEL_W_MM * mm
LABEL_H = LABEL_H_MM * mm

def _config_path():
    """En un .exe compilado la config se guarda junto al ejecutable."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "config.json"
    return Path(__file__).parent / "config.json"

CONFIG_PATH = _config_path()

VERSION = "2.2.3"
GITHUB_REPO = "Lacho101/SW-ETIQUETAS-GA"
UPDATE_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

FUENTES_PDF = {
    "Helvetica": ("Helvetica-Bold", "Helvetica"),
    "Arial":     ("Helvetica-Bold", "Helvetica"),
    "Times":     ("Times-Bold",     "Times-Roman"),
    "Courier":   ("Courier-Bold",   "Courier"),
}


# ── Estilos (Tema Naranja) ────────────────────────────────────────────────────
STYLE = """
QMainWindow, QWidget#root { background: #F7F8FA; }
QGroupBox {
    font-size: 11px; font-weight: 600; color: #444;
    border: 1.5px solid #DDE1E8; border-radius: 6px;
    margin-top: 10px; padding-top: 6px; background: #FFFFFF;
}
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 6px; left: 10px; }
QLabel { color: #333; font-size: 12px; }
QLineEdit, QTextEdit, QSpinBox, QComboBox, QTableWidget {
    background: #FFFFFF; border: 1.5px solid #DDE1E8; border-radius: 5px;
    padding: 4px 8px; font-size: 12px; color: #222;
    selection-background-color: #F29C2E;
}
QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, QComboBox:focus, QTableWidget:focus { border-color: #F08800; }
QComboBox::drop-down { border: none; }
QComboBox::down-arrow { image: none; width: 0; }
QPushButton#primary {
    background: #F08800; color: #fff; border: none; border-radius: 6px;
    font-size: 13px; font-weight: 700; padding: 9px 24px;
}
QPushButton#primary:hover { background: #D97A00; }
QPushButton#primary:pressed { background: #B36500; }
QPushButton#secondary {
    background: #EAECF2; color: #444; border: none; border-radius: 6px;
    font-size: 12px; padding: 9px 18px;
}
QPushButton#secondary:hover { background: #D5D9E4; }
QPushButton#danger {
    background: #FEF4E6; color: #B36500; border: 1.5px solid #F7C483;
    border-radius: 6px; font-size: 12px; padding: 9px 16px;
}
QPushButton#danger:hover { background: #FCE6CA; }
QRadioButton, QCheckBox { font-size: 12px; color: #333; spacing: 6px; }
QRadioButton::indicator, QCheckBox::indicator { width: 15px; height: 15px; }
QStatusBar { background: #EAECF2; color: #666; font-size: 11px; border-top: 1px solid #DDE1E8; }
"""

# ── Config ────────────────────────────────────────────────────────────────────
DEFAULT_CFG = {
    "fuente": "Helvetica",
    "tam_linea1": 7,
    "tam_linea2": 12,
    "copias": 1,
    "abrir_pdf": True,
    "auto_update": True,
    "modo": "manual",
    "mostrar_desc": True,
    "mostrar_cod": True,
    "mostrar_precio": True,
    "impresora": "",
    "serie": {
        "pref": "A", "n2i": "1", "n2f": "1",
        "n3i": "1", "n3f": "6",
        "n4i": "1", "n4f": "5",
    },
}

def resolver_ruta(ruta_relativa):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

def cargar_cfg():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH) as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CFG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return DEFAULT_CFG.copy()

def guardar_cfg(cfg: dict):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ── Actualizaciones automáticas ───────────────────────────────────────────────
def _es_frozen():
    return bool(getattr(sys, "frozen", False))

def _a_pares(v):
    """Convierte 'v2.1' o '2.1.0' en [2, 1, 0] para comparar versiones."""
    res = []
    for p in str(v).lstrip("v").replace("-", ".").split("."):
        try:
            res.append(int(p))
        except ValueError:
            res.append(0)
    return res

def _es_mas_nueva(actual, nueva):
    a, b = _a_pares(actual), _a_pares(nueva)
    n = max(len(a), len(b))
    for i in range(n):
        x = a[i] if i < len(a) else 0
        y = b[i] if i < len(b) else 0
        if y != x:
            return y > x
    return False

def _obtener_ultima_publicacion():
    """Consulta la última release del repo en GitHub."""
    req = urllib.request.Request(
        UPDATE_API,
        headers={"User-Agent": "SW-ETIQUETAS-GA", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req, timeout=10) as r:
        data = json.loads(r.read().decode("utf-8"))
    tag = str(data.get("tag_name", "")).lstrip("v")
    assets = data.get("assets", []) or []
    asset = next((a for a in assets if str(a.get("name", "")).lower().endswith(".exe")), None)
    return {
        "version": tag,
        "notas": data.get("body") or "",
        "prerelease": bool(data.get("prerelease")),
        "download_url": asset.get("browser_download_url") if asset else None,
        "asset_name": asset.get("name") if asset else None,
    }

class CheckUpdateWorker(QObject):
    listo = pyqtSignal(dict)
    fallo = pyqtSignal(str)

    def run(self):
        try:
            self.listo.emit(_obtener_ultima_publicacion())
        except Exception as ex:
            self.fallo.emit(str(ex))

class DescargaWorker(QObject):
    progreso = pyqtSignal(int, int)
    listo = pyqtSignal(str)
    fallo = pyqtSignal(str)

    def __init__(self, url, destino):
        super().__init__()
        self.url = url
        self.destino = destino

    def run(self):
        try:
            def hook(done, block, total):
                self.progreso.emit(done * block, total)
            urllib.request.urlretrieve(self.url, self.destino, reporthook=hook)
            self.listo.emit(self.destino)
        except Exception as ex:
            self.fallo.emit(str(ex))


# ── PDF ───────────────────────────────────────────────────────────────────────
def generar_pdf(etiquetas, ruta, copies=1, fuente="Helvetica", tam1=7, tam2=12, flags=None):
    if flags is None: flags = {"desc": True, "cod": True, "precio": True}
    bold, normal = FUENTES_PDF.get(fuente, ("Helvetica-Bold", "Helvetica"))
    c = rl_canvas.Canvas(ruta, pagesize=(LABEL_W, LABEL_H))
    for et in etiquetas:
        for _ in range(copies):
            _dibujar(c, et, bold, normal, tam1, tam2, flags)
            c.showPage()
    c.save()


def _dibujar(c, et, bold, normal, tam1, tam2, flags):
    """Dibuja una etiqueta centrando verticalmente todos los elementos activos."""
    desc = str(et.get("descripcion", "")).strip()
    cod  = str(et.get("codigo",      "")).strip()
    precio = str(et.get("precio",    "")).strip()

    has_desc   = flags["desc"]   and bool(desc)
    has_cod    = flags["cod"]    and bool(cod)
    has_precio = flags["precio"] and bool(precio)

    if not (has_desc or has_cod or has_precio):
        return

    cx = LABEL_W / 2
    MARGIN = 1.5 * mm          # margen vertical interior de la etiqueta
    GAP    = 1.2 * mm          # espacio entre bloques

    # ── Alturas de cada bloque ────────────────────────────────────────────────
    h_desc   = (tam1 / 72 * 25.4 * mm) if has_desc   else 0   # puntos → mm → pts
    h_bc     = 7.5 * mm        if has_cod    else 0            # altura barra
    h_cod_txt= (tam1 / 72 * 25.4 * mm) if has_cod    else 0   # texto bajo barra
    h_precio = (tam2 / 72 * 25.4 * mm) if has_precio else 0

    # Altura total del bloque de contenido
    bloques = []
    if has_desc:   bloques.append(h_desc)
    if has_cod:    bloques.append(h_bc + GAP * 0.5 + h_cod_txt)
    if has_precio: bloques.append(h_precio)

    total_h = sum(bloques) + GAP * (len(bloques) - 1)

    # ── Posición de inicio (top del bloque, centrado) ────────────────────────
    y_top = LABEL_H / 2 + total_h / 2   # ReportLab: y crece hacia arriba
    cursor = y_top

    if has_desc:
        c.setFont(normal, tam1)
        if len(desc) > 30: desc = desc[:28] + ".."
        # drawCentredString en ReportLab dibuja sobre la línea base
        c.drawCentredString(cx, cursor - h_desc * 0.8, desc)
        cursor -= h_desc + GAP

    if has_cod:
        try:
            # Ancho máximo disponible restando los márgenes horizontales
            margen_h = 1.5 * mm
            avail_w = LABEL_W - 2 * margen_h
            bc = code128.Code128(cod, barHeight=h_bc, barWidth=1.2)
            lq, rq = bc.lquiet, bc.rquiet
            quiet_total = (lq + rq) if bc.quiet else 0
            bars_w = bc.width - quiet_total          # ancho solo de barras
            available = avail_w - quiet_total        # espacio para barras
            # Ajusta el grosor de las barras para que el código nunca se salga
            bar_w = min(1.2, 1.2 * (available / bars_w)) if bars_w > 0 else 1.2
            bc = code128.Code128(cod, barHeight=h_bc, barWidth=bar_w)
            bc.drawOn(c, cx - bc.width / 2.0, cursor - h_bc)
            c.setFont(bold, tam1)
            c.drawCentredString(cx, cursor - h_bc - h_cod_txt * 0.9, cod)
        except Exception:
            c.setFont(bold, tam2)
            c.drawCentredString(cx, cursor - h_bc, f"ERR: {cod}")
        cursor -= (h_bc + h_cod_txt + GAP)

    if has_precio:
        ptxt = f"${precio}" if not precio.startswith("$") else precio
        c.setFont(bold, tam2)
        c.drawCentredString(cx, cursor - h_precio * 0.8, ptxt)


def get_printers() -> list[str]:
    """Devuelve la lista de impresoras instaladas en Windows."""
    try:
        result = subprocess.run(
            ["wmic", "printer", "get", "name"],
            capture_output=True, text=True, timeout=5
        )
        lines = [l.strip() for l in result.stdout.splitlines() if l.strip() and l.strip().lower() != "name"]
        return lines
    except Exception:
        return []


def get_default_printer() -> str:
    """Devuelve el nombre de la impresora predeterminada del sistema."""
    try:
        result = subprocess.run(
            ["wmic", "printer", "where", "Default=TRUE", "get", "name"],
            capture_output=True, text=True, timeout=5
        )
        lines = [l.strip() for l in result.stdout.splitlines() if l.strip() and l.strip().lower() != "name"]
        return lines[0] if lines else ""
    except Exception:
        return ""


def imprimir_pdf_windows(ruta: str, impresora: str = ""):
    """Envía el PDF a la impresora especificada (o a la predeterminada si está vacío)."""
    if impresora:
        # Usar SumatraPDF si está disponible para seleccionar impresora específica
        sumatra_paths = [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        ]
        sumatra = next((p for p in sumatra_paths if os.path.exists(p)), None)
        if sumatra:
            subprocess.Popen([
                sumatra, "-print-to", impresora,
                "-print-settings", "fit",
                "-silent", ruta
            ])
            return
        # Fallback: usar PowerShell para enviar a impresora específica
        try:
            script = (
                f'$pdf = "{ruta}"; '
                f'$printer = "{impresora}"; '
                f'$shell = New-Object -ComObject Shell.Application; '
                f'$folder = $shell.Namespace((Split-Path $pdf)); '
                f'$file = $folder.ParseName((Split-Path $pdf -Leaf)); '
                f'$file.InvokeVerbEx("printto", $printer)'
            )
            subprocess.Popen(
                ["powershell", "-WindowStyle", "Hidden", "-Command", script]
            )
            return
        except Exception:
            pass
    # Comportamiento original: impresora predeterminada del sistema
    import ctypes
    SW_HIDE = 0
    ctypes.windll.shell32.ShellExecuteW(
        None, "print", ruta, None, None, SW_HIDE
    )


def abrir_pdf(ruta: str):
    # Intentar abrir con Brave primero
    brave_paths = [
        r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
        r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\Application\brave.exe"),
    ]
    for bpath in brave_paths:
        if os.path.exists(bpath):
            try:
                subprocess.Popen([bpath, ruta])
                return
            except Exception:
                pass
    # Fallback: abrir con el navegador/visor predeterminado
    try:
        os.startfile(ruta)
    except Exception:
        pass


# ── Vista previa de etiqueta ──────────────────────────────────────────────────
class EtiquetaPreview(QWidget):
    SCALE = 4.0

    def __init__(self, parent=None):
        super().__init__(parent)
        pw = int(LABEL_W_MM * self.SCALE)
        ph = int(LABEL_H_MM * self.SCALE)
        self.setFixedSize(pw + 20, ph + 20)
        self._data = {}
        self._flags = {"desc": True, "cod": True, "precio": True}
        self._fuente = "Arial"
        self._tam1 = 7
        self._tam2 = 12

    def set_data(self, data, flags, fuente, tam1, tam2):
        self._data = data
        self._flags = flags
        self._fuente = fuente
        self._tam1 = tam1
        self._tam2 = tam2
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        pw = int(LABEL_W_MM * self.SCALE)
        ph = int(LABEL_H_MM * self.SCALE)
        ox, oy = 10, 10

        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QBrush(QColor("#D0D4DC")))
        p.drawRoundedRect(ox + 3, oy + 3, pw, ph, 4, 4)

        p.setBrush(QBrush(QColor("#FFFFFF")))
        p.setPen(QPen(QColor("#BBCCE8"), 1.5))
        p.drawRoundedRect(ox, oy, pw, ph, 4, 4)

        if not self._data:
            p.end()
            return

        cx = ox + pw / 2
        cy = oy + ph / 2

        font_family = self._fuente if self._fuente in ("Arial", "Times New Roman", "Courier New") else "Arial"
        if self._fuente == "Helvetica": font_family = "Arial"
        elif self._fuente == "Times": font_family = "Times New Roman"
        elif self._fuente == "Courier": font_family = "Courier New"
        
        desc = str(self._data.get("descripcion", "")).strip()
        cod = str(self._data.get("codigo", "")).strip()
        precio = str(self._data.get("precio", "")).strip()

        has_desc   = self._flags["desc"]   and bool(desc)
        has_cod    = self._flags["cod"]    and bool(cod)
        has_precio = self._flags["precio"] and bool(precio)

        if not (has_desc or has_cod or has_precio):
            p.end()
            return

        MARGIN = int(ph * 0.04)    # margen interior vertical
        GAP    = int(ph * 0.06)    # espacio entre bloques

        # ── Fuentes y métricas ────────────────────────────────────────────────
        f_small = QFont(font_family, int(self._tam1 * self.SCALE * 0.35))
        f_bold  = QFont(font_family, int(self._tam1 * self.SCALE * 0.35), QFont.Weight.Bold)
        f_price = QFont(font_family, int(self._tam2 * self.SCALE * 0.35), QFont.Weight.Bold)
        fm_small = QFontMetrics(f_small)
        fm_bold  = QFontMetrics(f_bold)
        fm_price = QFontMetrics(f_price)

        bw = int(pw * 0.75)
        bh = int(ph * 0.32)  # altura de barras por defecto con otros elementos
        if has_cod and not has_desc and not has_precio:
            bh = int(ph * 0.50)  # solo código → barras más altas

        # ── Calcular altura total del bloque de contenido ─────────────────────
        bloques_h = []
        if has_desc:   bloques_h.append(fm_small.height())
        if has_cod:    bloques_h.append(bh + int(GAP * 0.4) + fm_bold.height())
        if has_precio: bloques_h.append(fm_price.height())

        total_h = sum(bloques_h) + GAP * (len(bloques_h) - 1)

        # ── Punto de inicio (top del bloque, centrado en etiqueta) ────────────
        y_cursor = int(oy + (ph - total_h) / 2)  # top del primer bloque

        if has_desc:
            p.setFont(f_small)
            p.setPen(QColor("#555"))
            tdesc = desc if len(desc) <= 30 else desc[:28] + ".."
            tx = int(cx - fm_small.horizontalAdvance(tdesc) / 2)
            ty = y_cursor + fm_small.ascent()
            p.drawText(tx, ty, tdesc)
            y_cursor += fm_small.height() + GAP

        if has_cod:
            # Barras
            p.setBrush(QBrush(QColor("#000")))
            p.setPen(Qt.PenStyle.NoPen)
            bar_x = int(cx - bw / 2)
            for i in range(0, bw, 6):
                w = 2 if i % 4 == 0 else 4
                p.drawRect(bar_x + i, y_cursor, w, bh)
            y_cursor += bh + int(GAP * 0.4)
            # Texto del código
            p.setFont(f_bold)
            p.setPen(QColor("#111"))
            tx = int(cx - fm_bold.horizontalAdvance(cod) / 2)
            ty = y_cursor + fm_bold.ascent()
            p.drawText(tx, ty, cod)
            y_cursor += fm_bold.height() + GAP

        if has_precio:
            ptxt = f"${precio}" if not precio.startswith("$") else precio
            p.setFont(f_price)
            p.setPen(QColor("#111"))
            tx = int(cx - fm_price.horizontalAdvance(ptxt) / 2)
            ty = y_cursor + fm_price.ascent()
            p.drawText(tx, ty, ptxt)

        p.end()


# ── Ventana principal ─────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.cfg = cargar_cfg()
        self.setWindowTitle("Generador Etiquetas GA")
        self.setWindowIcon(QIcon(resolver_ruta("icono.ico")))
        self.setMinimumSize(950, 680)
        self.setObjectName("root")
        
        self.excel_path = ""
        self.excel_data = []
        
        self._build_ui()
        self._cargar_cfg_ui()
        self._actualizar_preview()
        QTimer.singleShot(2500, self._comprobar_actualizacion)

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("root")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)
        root.setContentsMargins(14, 12, 14, 10)

        root.addWidget(self._header())

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(8)

        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setSpacing(10)
        lv.setContentsMargins(0, 0, 4, 0)
        
        lv.addWidget(self._modo_group())
        
        # Paneles de entrada
        self.grp_manual = self._input_manual_group()
        self.grp_serie = self._input_serie_group()
        self.grp_excel = self._input_excel_group()
        
        lv.addWidget(self.grp_manual, 1)
        lv.addWidget(self.grp_serie, 1)
        lv.addWidget(self.grp_excel, 1)
        
        lv.addWidget(self._opciones_group())
        lv.addWidget(self._botones_frame())

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setSpacing(10)
        rv.setContentsMargins(4, 0, 0, 0)
        rv.addWidget(self._preview_group())

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([600, 320])
        root.addWidget(splitter, 1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Listo.")

    def _header(self):
        frm = QFrame()
        frm.setStyleSheet("background: #F6A23B; border-radius: 7px;")
        hl = QHBoxLayout(frm)
        hl.setContentsMargins(16, 10, 16, 10)
        title = QLabel("GRAN ALMACÉN ETIQUETAS                   by Lacho")
        title.setStyleSheet("color:#fff; font-size:15px; font-weight:700; background:transparent;")
        size_lbl = QLabel(f"51.20 × 25.00 mm  |  EVA58  |  v{VERSION}")
        size_lbl.setStyleSheet("color:#FFFFFF; font-size:11px; background:transparent;")
        self.btn_update_hdr = QPushButton("Actualizar")
        self.btn_update_hdr.setObjectName("secondary")
        self.btn_update_hdr.setStyleSheet(
            "QPushButton { background:#FFFFFF; color:#F08800; border:none; border-radius:4px; "
            "font-size:11px; font-weight:700; padding:4px 10px; } "
            "QPushButton:hover { background:#FFF3E0; }"
        )
        self.btn_update_hdr.setVisible(False)
        self.btn_update_hdr.clicked.connect(self._actualizar_desde_header)
        hl.addWidget(title)
        hl.addStretch()
        hl.addWidget(size_lbl)
        hl.addWidget(self.btn_update_hdr)
        return frm

    def _modo_group(self):
        grp = QGroupBox("Modo de entrada")
        hl = QHBoxLayout(grp)
        hl.setSpacing(20)
        self.rb_manual = QRadioButton("Manual")
        self.rb_serie  = QRadioButton("Serie")
        self.rb_excel  = QRadioButton("Desde Excel")
        self.btn_grp   = QButtonGroup()
        self.btn_grp.addButton(self.rb_manual, 0)
        self.btn_grp.addButton(self.rb_serie, 1)
        self.btn_grp.addButton(self.rb_excel, 2)
        
        self.rb_manual.setChecked(True)
        self.btn_grp.buttonClicked.connect(self._on_modo)
        
        hl.addWidget(self.rb_manual)
        hl.addWidget(self.rb_serie)
        hl.addWidget(self.rb_excel)
        hl.addStretch()
        return grp

    def _input_manual_group(self):
        grp = QGroupBox("Etiquetas Manuales")
        vl = QVBoxLayout(grp)
        self.table_manual = QTableWidget(1, 3)
        self.table_manual.setHorizontalHeaderLabels(["Descripción", "Código", "Precio"])
        self.table_manual.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_manual.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_manual.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_manual.itemChanged.connect(self._actualizar_preview)
        
        btn_hl = QHBoxLayout()
        btn_add = QPushButton("+ Fila")
        btn_add.setObjectName("secondary")
        btn_add.clicked.connect(lambda: self.table_manual.insertRow(self.table_manual.rowCount()))
        btn_rm = QPushButton("- Fila")
        btn_rm.setObjectName("secondary")
        btn_rm.clicked.connect(lambda: self.table_manual.removeRow(self.table_manual.currentRow()) if self.table_manual.rowCount() > 1 else None)
        btn_hl.addWidget(btn_add)
        btn_hl.addWidget(btn_rm)
        btn_hl.addStretch()
        
        vl.addWidget(self.table_manual)
        vl.addLayout(btn_hl)
        return grp

    def _input_serie_group(self):
        grp = QGroupBox("Parámetros de serie")
        grid = QGridLayout(grp)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        campos = [
            ("Prefijo:", "pref", 0, 0),
            ("N2 inicio:", "n2i", 0, 2), ("N2 fin:", "n2f", 0, 4),
            ("N3 inicio:", "n3i", 1, 0), ("N3 fin:", "n3f", 1, 2),
            ("N4 inicio:", "n4i", 1, 4), ("N4 fin:", "n4f", 1, 6),
        ]
        self._serie_vars = {}
        for lbl_txt, key, row, col in campos:
            grid.addWidget(QLabel(lbl_txt), row, col)
            le = QLineEdit()
            le.setFixedWidth(50)
            le.textChanged.connect(self._actualizar_preview)
            self._serie_vars[key] = le
            grid.addWidget(le, row, col + 1)

        btn_regen = QPushButton("↻ Regenerar")
        btn_regen.setObjectName("secondary")
        btn_regen.clicked.connect(self._actualizar_preview)
        grid.addWidget(btn_regen, 2, 0, 1, 4)
        grp.hide()
        return grp

    def _input_excel_group(self):
        grp = QGroupBox("Carga desde Excel")
        vl = QVBoxLayout(grp)
        
        hl = QHBoxLayout()
        self.btn_excel = QPushButton("📁 Seleccionar Archivo")
        self.btn_excel.setObjectName("secondary")
        self.btn_excel.clicked.connect(self._cargar_excel)
        
        self.lbl_excel_path = QLabel("Ningún archivo seleccionado")
        self.lbl_excel_path.setWordWrap(True)
        
        hl.addWidget(self.btn_excel)
        hl.addWidget(self.lbl_excel_path, 1)
        vl.addLayout(hl)

        # Controles de mapeo de columnas
        self.frm_mapping = QFrame()
        gl = QGridLayout(self.frm_mapping)
        gl.setContentsMargins(0, 5, 0, 0)
        
        gl.addWidget(QLabel("Col. Descripción:"), 0, 0)
        self.cb_col_desc = QComboBox()
        self.cb_col_desc.currentIndexChanged.connect(self._procesar_datos_excel)
        gl.addWidget(self.cb_col_desc, 0, 1)

        gl.addWidget(QLabel("Col. Código:"), 1, 0)
        self.cb_col_cod = QComboBox()
        self.cb_col_cod.currentIndexChanged.connect(self._procesar_datos_excel)
        gl.addWidget(self.cb_col_cod, 1, 1)

        gl.addWidget(QLabel("Col. Precio:"), 2, 0)
        self.cb_col_precio = QComboBox()
        self.cb_col_precio.currentIndexChanged.connect(self._procesar_datos_excel)
        gl.addWidget(self.cb_col_precio, 2, 1)

        gl.addWidget(QLabel("Col. Cantidad:"), 3, 0)
        self.cb_col_cantidad = QComboBox()
        self.cb_col_cantidad.setToolTip("Nº de etiquetas a imprimir para cada fila (si está vacío o no se selecciona, se imprime 1).")
        self.cb_col_cantidad.currentIndexChanged.connect(self._procesar_datos_excel)
        gl.addWidget(self.cb_col_cantidad, 3, 1)

        self.frm_mapping.setEnabled(False)
        vl.addWidget(self.frm_mapping)
        
        vl.addStretch()
        grp.hide()
        return grp

    def _opciones_group(self):
        grp = QGroupBox("Opciones e Impresión")
        gl = QGridLayout(grp)
        gl.setHorizontalSpacing(12)
        gl.setVerticalSpacing(6)

        # Opciones de visibilidad
        self.chk_desc = QCheckBox("Imprimir Descripción")
        self.chk_cod = QCheckBox("Imprimir Código")
        self.chk_precio = QCheckBox("Imprimir Precio")
        self.chk_desc.stateChanged.connect(self._actualizar_preview)
        self.chk_cod.stateChanged.connect(self._actualizar_preview)
        self.chk_precio.stateChanged.connect(self._actualizar_preview)
        
        gl.addWidget(self.chk_desc, 0, 0)
        gl.addWidget(self.chk_cod, 0, 1)
        gl.addWidget(self.chk_precio, 0, 2)

        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setStyleSheet("color:#eee")
        gl.addWidget(line, 1, 0, 1, 6)

        # Opciones de fuente
        gl.addWidget(QLabel("Fuente:"), 2, 0)
        self.cb_fuente = QComboBox()
        self.cb_fuente.addItems(["Helvetica", "Arial", "Times", "Courier"])
        self.cb_fuente.currentTextChanged.connect(self._actualizar_preview)
        gl.addWidget(self.cb_fuente, 2, 1)

        gl.addWidget(QLabel("Tam 1:"), 2, 2)
        self.sb_tam1 = QSpinBox()
        self.sb_tam1.setRange(5, 20); self.sb_tam1.setValue(7)
        self.sb_tam1.valueChanged.connect(self._actualizar_preview)
        gl.addWidget(self.sb_tam1, 2, 3)

        gl.addWidget(QLabel("Tam 2:"), 2, 4)
        self.sb_tam2 = QSpinBox()
        self.sb_tam2.setRange(6, 36); self.sb_tam2.setValue(12)
        self.sb_tam2.valueChanged.connect(self._actualizar_preview)
        gl.addWidget(self.sb_tam2, 2, 5)

        gl.addWidget(QLabel("Copias/etiqueta:"), 3, 0)
        self.sb_copias = QSpinBox()
        self.sb_copias.setRange(1, 99); self.sb_copias.setValue(1)
        self.sb_copias.valueChanged.connect(self._actualizar_preview)
        gl.addWidget(self.sb_copias, 3, 1)

        self.chk_abrir = QCheckBox("Abrir PDF")
        self.chk_abrir.setChecked(True)
        gl.addWidget(self.chk_abrir, 3, 2)

        self.chk_update = QCheckBox("Buscar actualizaciones")
        gl.addWidget(self.chk_update, 3, 3)

        btn_upd = QPushButton("Buscar ahora")
        btn_upd.setObjectName("secondary")
        btn_upd.setToolTip("Comprobar si hay una versión más reciente en GitHub")
        btn_upd.clicked.connect(lambda: self._comprobar_actualizacion(forzar=True))
        gl.addWidget(btn_upd, 3, 4)

        # ── Selector de impresora ──────────────────────────────────────────────
        line2 = QFrame(); line2.setFrameShape(QFrame.Shape.HLine); line2.setStyleSheet("color:#eee")
        gl.addWidget(line2, 4, 0, 1, 6)

        gl.addWidget(QLabel("🖨  Impresora:"), 5, 0)
        self.cb_impresora = QComboBox()
        self.cb_impresora.setMinimumWidth(260)
        self.cb_impresora.setToolTip(
            "Selecciona la impresora a usar. '(Predeterminada del sistema)' usa la impresora por defecto de Windows."
        )
        self._recargar_impresoras()
        gl.addWidget(self.cb_impresora, 5, 1, 1, 4)

        btn_refresh = QPushButton("↻")
        btn_refresh.setObjectName("secondary")
        btn_refresh.setFixedWidth(32)
        btn_refresh.setToolTip("Actualizar lista de impresoras")
        btn_refresh.clicked.connect(self._recargar_impresoras)
        gl.addWidget(btn_refresh, 5, 5)

        return grp

    def _recargar_impresoras(self):
        """Recarga la lista de impresoras disponibles en el sistema."""
        impresora_actual = self.cb_impresora.currentText() if hasattr(self, "cb_impresora") else ""
        printers = get_printers()
        default_name = get_default_printer()

        if hasattr(self, "cb_impresora"):
            self.cb_impresora.blockSignals(True)
            self.cb_impresora.clear()

        # Primera opción: predeterminada del sistema
        default_label = f"(Predeterminada del sistema)"
        if default_name:
            default_label = f"(Predeterminada)  {default_name}"

        if hasattr(self, "cb_impresora"):
            self.cb_impresora.addItem(default_label)
            for p in printers:
                self.cb_impresora.addItem(p)

            # Restaurar selección previa si sigue en la lista
            idx = self.cb_impresora.findText(impresora_actual)
            if idx >= 0:
                self.cb_impresora.setCurrentIndex(idx)

            self.cb_impresora.blockSignals(False)

    def _botones_frame(self):
        frm = QFrame()
        hl = QHBoxLayout(frm)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)

        self.lbl_count = QLabel("")
        self.lbl_count.setStyleSheet("color:#888; font-size:11px;")

        btn_pdf = QPushButton("Guardar PDF")
        btn_pdf.setObjectName("primary")
        btn_pdf.clicked.connect(self._accion_pdf)

        btn_print = QPushButton("Imprimir")
        btn_print.setObjectName("primary")
        btn_print.setStyleSheet(
            "QPushButton { background:#28A745; color:#fff; border:none; "
            "border-radius:6px; font-size:13px; font-weight:700; padding:9px 20px; }"
            "QPushButton:hover { background:#1E8035; }"
        )
        btn_print.clicked.connect(self._accion_imprimir)

        btn_limpiar = QPushButton("Limpiar")
        btn_limpiar.setObjectName("danger")
        btn_limpiar.clicked.connect(self._limpiar)

        hl.addWidget(btn_pdf)
        hl.addWidget(btn_print)
        hl.addWidget(btn_limpiar)
        hl.addStretch()
        hl.addWidget(self.lbl_count)
        return frm

    def _preview_group(self):
        grp = QGroupBox("Vista previa")
        vl = QVBoxLayout(grp)
        vl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)

        lbl = QLabel("Primera etiqueta de la lista:")
        lbl.setStyleSheet("color:#888; font-size:11px;")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        vl.addWidget(lbl)

        self.preview_widget = EtiquetaPreview()
        vl.addWidget(self.preview_widget, 0, Qt.AlignmentFlag.AlignHCenter)

        self.lbl_preview_txt = QLabel("")
        self.lbl_preview_txt.setWordWrap(True)
        self.lbl_preview_txt.setStyleSheet(
            "background:#F0F4FF; border:1px solid #C8D8F8; border-radius:5px; "
            "padding:6px; font-size:11px; color:#333;"
        )
        self.lbl_preview_txt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        vl.addWidget(self.lbl_preview_txt)
        vl.addStretch()
        return grp

    # ── Lógica ────────────────────────────────────────────────────────────────
    def _on_modo(self):
        self.grp_manual.hide()
        self.grp_serie.hide()
        self.grp_excel.hide()
        
        if self.rb_manual.isChecked():
            self.grp_manual.show()
        elif self.rb_serie.isChecked():
            self.grp_serie.show()
        else:
            self.grp_excel.show()
            
        self._actualizar_preview()

    def _get_flags(self):
        return {
            "desc": self.chk_desc.isChecked(),
            "cod": self.chk_cod.isChecked(),
            "precio": self.chk_precio.isChecked()
        }

    def _get_etiquetas(self):
        if self.rb_manual.isChecked():
            res = []
            for r in range(self.table_manual.rowCount()):
                id1 = self.table_manual.item(r, 0)
                id2 = self.table_manual.item(r, 1)
                id3 = self.table_manual.item(r, 2)
                d = id1.text() if id1 else ""
                c = id2.text() if id2 else ""
                p = id3.text() if id3 else ""
                if d or c or p:
                    res.append({"descripcion": d, "codigo": c, "precio": p})
            return res
        elif self.rb_serie.isChecked():
            return self._calcular_serie()
        else:
            res = []
            for e in self.excel_data:
                cant = e.get("cantidad", 1)
                for _ in range(cant):
                    res.append({"descripcion": e.get("descripcion", ""),
                                "codigo": e.get("codigo", ""),
                                "precio": e.get("precio", "")})
            return res

    def _calcular_serie(self):
        try:
            pref = self._serie_vars["pref"].text().strip() or "A"
            n2i  = int(self._serie_vars["n2i"].text())
            n2f  = int(self._serie_vars["n2f"].text())
            n3i  = int(self._serie_vars["n3i"].text())
            n3f  = int(self._serie_vars["n3f"].text())
            n4i  = int(self._serie_vars["n4i"].text())
            n4f  = int(self._serie_vars["n4f"].text())
        except ValueError:
            return []
        
        res = []
        for n2 in range(n2i, n2f + 1):
            for n3 in range(n3i, n3f + 1):
                for n4 in range(n4i, n4f + 1):
                    cod = f"{pref}.{n2}.{n3}.{n4}"
                    res.append({"descripcion": "Serie GA", "codigo": cod, "precio": ""})
        return res

    def _cargar_excel(self):
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not ruta:
            return
            
        try:
            self.status_bar.showMessage("Leyendo cabeceras de Excel...")
            QApplication.processEvents()
            
            self.excel_wb = openpyxl.load_workbook(ruta, data_only=True)
            self.excel_ws = self.excel_wb.active
            
            self.excel_headers = []
            for cell in self.excel_ws[1]:
                val = str(cell.value).strip() if cell.value else f"Col {cell.column}"
                self.excel_headers.append(val)
                
            self.cb_col_desc.blockSignals(True)
            self.cb_col_cod.blockSignals(True)
            self.cb_col_precio.blockSignals(True)
            self.cb_col_cantidad.blockSignals(True)

            for cb in (self.cb_col_desc, self.cb_col_cod, self.cb_col_precio, self.cb_col_cantidad):
                cb.clear()
                cb.addItem("-- Ninguna --")
                cb.addItems(self.excel_headers)

            for i, h in enumerate(self.excel_headers):
                h_low = h.lower()
                if "desc" in h_low: self.cb_col_desc.setCurrentIndex(i + 1)
                elif "cód" in h_low or "cod" in h_low: self.cb_col_cod.setCurrentIndex(i + 1)
                elif "pre" in h_low: self.cb_col_precio.setCurrentIndex(i + 1)
                elif "cant" in h_low or "nº" in h_low or "num" in h_low: self.cb_col_cantidad.setCurrentIndex(i + 1)

            self.cb_col_desc.blockSignals(False)
            self.cb_col_cod.blockSignals(False)
            self.cb_col_precio.blockSignals(False)
            self.cb_col_cantidad.blockSignals(False)

            self.excel_path = ruta
            self.lbl_excel_path.setText(os.path.basename(ruta))
            self.frm_mapping.setEnabled(True)
            
            self._procesar_datos_excel()
            
        except Exception as ex:
            QMessageBox.critical(self, "Error Excel", f"No se pudo leer el archivo:\n{str(ex)}")
            self.status_bar.showMessage("Error al leer Excel")

    def _procesar_datos_excel(self):
        if not hasattr(self, 'excel_ws') or not self.excel_ws:
            return
            
        idx_desc = self.cb_col_desc.currentIndex() - 1
        idx_cod = self.cb_col_cod.currentIndex() - 1
        idx_precio = self.cb_col_precio.currentIndex() - 1
        idx_cant = self.cb_col_cantidad.currentIndex() - 1

        self.excel_data = []
        for row in self.excel_ws.iter_rows(min_row=2, values_only=True):
            d = str(row[idx_desc]) if idx_desc >= 0 and idx_desc < len(row) and row[idx_desc] is not None else ""
            c = str(row[idx_cod]) if idx_cod >= 0 and idx_cod < len(row) and row[idx_cod] is not None else ""
            p = str(row[idx_precio]) if idx_precio >= 0 and idx_precio < len(row) and row[idx_precio] is not None else ""

            cant = 1
            if idx_cant >= 0 and idx_cant < len(row) and row[idx_cant] is not None:
                try:
                    cant = int(float(str(row[idx_cant]).replace(",", ".")))
                except (ValueError, TypeError):
                    cant = 1
                if cant < 1:
                    cant = 1

            if d or c or p:
                self.excel_data.append({"descripcion": d, "codigo": c, "precio": p, "cantidad": cant})

        total = sum(e.get("cantidad", 1) for e in self.excel_data)
        self.status_bar.showMessage(
            f"Excel cargado: {len(self.excel_data)} filas, {total} etiquetas"
        )
        self._actualizar_preview()

    def _actualizar_preview(self):
        etiquetas = self._get_etiquetas()
        n = len(etiquetas)
        copias = self.sb_copias.value() if hasattr(self, "sb_copias") else 1
        fuente = self.cb_fuente.currentText() if hasattr(self, "cb_fuente") else "Arial"
        tam1 = self.sb_tam1.value() if hasattr(self, "sb_tam1") else 7
        tam2 = self.sb_tam2.value() if hasattr(self, "sb_tam2") else 12
        flags = self._get_flags()

        if n == 0:
            self.preview_widget.set_data({}, flags, fuente, tam1, tam2)
            self.lbl_preview_txt.setText("Sin etiquetas")
            self.lbl_count.setText("")
            return

        self.preview_widget.set_data(etiquetas[0], flags, fuente, tam1, tam2)

        muestra = " | ".join([f"{e.get('codigo','')} ({e.get('precio','')})" for e in etiquetas[:4]])
        if n > 4:
            muestra += f"  … +{n-4} más"
        self.lbl_preview_txt.setText(muestra)

        total = n * copias
        self.lbl_count.setText(
            f"{n} etiqueta{'s' if n!=1 else ''}  ×  {copias} copia{'s' if copias!=1 else ''}  =  {total} páginas"
        )

    def _limpiar(self):
        if self.rb_manual.isChecked():
            self.table_manual.setRowCount(0)
            self.table_manual.insertRow(0)
        elif self.rb_excel.isChecked():
            self.excel_data = []
            self.excel_path = ""
            self.lbl_excel_path.setText("Ningún archivo seleccionado")
        self._actualizar_preview()

    def _datos_validos(self):
        et = self._get_etiquetas()
        if not et:
            QMessageBox.warning(self, "Sin datos", "No hay etiquetas para generar.")
            return None
        flags = self._get_flags()
        if not any(flags.values()):
            QMessageBox.warning(self, "Sin contenido", "Seleccione al menos un elemento para imprimir (Descripción, Código o Precio).")
            return None
        return et

    def _accion_pdf(self):
        etiquetas = self._datos_validos()
        if etiquetas is None:
            return
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar PDF", "etiquetas_ubicacion.pdf", "PDF (*.pdf)"
        )
        if not ruta:
            return
        self._generar_y_guardar(etiquetas, ruta, abrir=self.chk_abrir.isChecked())

    def _accion_imprimir(self):
        etiquetas = self._datos_validos()
        if etiquetas is None:
            return
        import tempfile
        tmp = tempfile.mktemp(suffix=".pdf")
        ok = self._generar_y_guardar(etiquetas, tmp, abrir=False, silencioso=True)
        if ok:
            try:
                # Determinar impresora seleccionada
                impresora_sel = self.cb_impresora.currentText() if hasattr(self, "cb_impresora") else ""
                # Si es la opción predeterminada (empieza con "("), no pasamos nombre
                if impresora_sel.startswith("("):
                    impresora_sel = ""
                imprimir_pdf_windows(tmp, impresora=impresora_sel)
                nombre_imp = impresora_sel if impresora_sel else "predeterminada"
                self.status_bar.showMessage(
                    f"✔  Enviado a impresora [{nombre_imp}]: {len(etiquetas) * self.sb_copias.value()} páginas"
                )
            except Exception as ex:
                QMessageBox.critical(self, "Error de impresión", str(ex))

    def _generar_y_guardar(self, etiquetas, ruta, abrir=False, silencioso=False):
        copias = self.sb_copias.value()
        fuente = self.cb_fuente.currentText()
        tam1   = self.sb_tam1.value()
        tam2   = self.sb_tam2.value()
        flags  = self._get_flags()
        total  = len(etiquetas) * copias
        
        try:
            self.status_bar.showMessage(f"Generando {total} páginas…")
            QApplication.processEvents()
            generar_pdf(etiquetas, ruta, copies=copias,
                        fuente=fuente, tam1=tam1, tam2=tam2, flags=flags)
            self._guardar_cfg()
            self.status_bar.showMessage(
                f"✔  PDF guardado: {os.path.basename(ruta)}  ({total} páginas)"
            )
            if not silencioso:
                QMessageBox.information(
                    self, "PDF generado",
                    f"Etiquetas: {len(etiquetas)}\n"
                    f"Copias c/u: {copias}\n"
                    f"Total páginas: {total}\n\n{ruta}"
                )
            if abrir:
                abrir_pdf(ruta)
            return True
        except Exception as ex:
            self.status_bar.showMessage("Error al generar el PDF.")
            QMessageBox.critical(self, "Error", str(ex))
            return False

    # ── Config ────────────────────────────────────────────────────────────────
    def _cargar_cfg_ui(self):
        cfg = self.cfg
        if cfg.get("modo") == "serie":
            self.rb_serie.setChecked(True)
        elif cfg.get("modo") == "excel":
            self.rb_excel.setChecked(True)
        else:
            self.rb_manual.setChecked(True)
            
        self._on_modo()

        idx = self.cb_fuente.findText(cfg.get("fuente", "Helvetica"))
        if idx >= 0: self.cb_fuente.setCurrentIndex(idx)

        self.sb_tam1.setValue(cfg.get("tam_linea1", 7))
        self.sb_tam2.setValue(cfg.get("tam_linea2", 12))
        self.sb_copias.setValue(cfg.get("copias", 1))
        self.chk_abrir.setChecked(cfg.get("abrir_pdf", True))
        self.chk_update.setChecked(cfg.get("auto_update", True))
        
        self.chk_desc.setChecked(cfg.get("mostrar_desc", True))
        self.chk_cod.setChecked(cfg.get("mostrar_cod", True))
        self.chk_precio.setChecked(cfg.get("mostrar_precio", True))

        for key, le in self._serie_vars.items():
            le.setText(cfg.get("serie", {}).get(key, DEFAULT_CFG["serie"][key]))

        # Restaurar impresora guardada
        impresora_guardada = cfg.get("impresora", "")
        if impresora_guardada and hasattr(self, "cb_impresora"):
            idx_imp = self.cb_impresora.findText(impresora_guardada)
            if idx_imp >= 0:
                self.cb_impresora.setCurrentIndex(idx_imp)

    def _guardar_cfg(self):
        modo = "manual"
        if self.rb_serie.isChecked(): modo = "serie"
        elif self.rb_excel.isChecked(): modo = "excel"

        # Guardar impresora seleccionada (vacío si es predeterminada del sistema)
        impresora_sel = self.cb_impresora.currentText() if hasattr(self, "cb_impresora") else ""
        if impresora_sel.startswith("("):
            impresora_sel = ""
        
        cfg = {
            "fuente": self.cb_fuente.currentText(),
            "tam_linea1": self.sb_tam1.value(),
            "tam_linea2": self.sb_tam2.value(),
            "copias": self.sb_copias.value(),
            "abrir_pdf": self.chk_abrir.isChecked(),
            "auto_update": self.chk_update.isChecked(),
            "modo": modo,
            "mostrar_desc": self.chk_desc.isChecked(),
            "mostrar_cod": self.chk_cod.isChecked(),
            "mostrar_precio": self.chk_precio.isChecked(),
            "impresora": impresora_sel,
            "serie": {k: v.text() for k, v in self._serie_vars.items()},
        }
        guardar_cfg(cfg)

    # ── Actualizaciones: hacer check/download/instalar ──────────────────────────
    def _comprobar_actualizacion(self, forzar=False):
        if not forzar and not (getattr(self, "chk_update", None) and self.chk_update.isChecked()):
            return
        if getattr(self, "_update_checking", False):
            return
        self._update_checking = True
        try:
            self.status_bar.showMessage("Comprobando actualizaciones…")
            self.thread_check = QThread(self)
            self.worker_check = CheckUpdateWorker()
            self.worker_check.moveToThread(self.thread_check)
            self.thread_check.started.connect(self.worker_check.run)
            self.thread_check.finished.connect(self.thread_check.deleteLater)
            self.worker_check.listo.connect(self._on_check_ok)
            self.worker_check.fallo.connect(self._on_check_fail)
            self.thread_check.start()
        except Exception:
            self._update_checking = False

    def _on_check_ok(self, data):
        self._update_checking = False
        try:
            self.thread_check.quit()
        except Exception:
            pass
        nueva = str(data.get("version", "")).lstrip("v")
        actual = str(VERSION).lstrip("v")
        if data.get("prerelease"):
            if hasattr(self, "btn_update_hdr"):
                self.btn_update_hdr.setVisible(False)
            return
        if not nueva or not _es_mas_nueva(actual, nueva):
            if hasattr(self, "btn_update_hdr"):
                self.btn_update_hdr.setVisible(False)
            if self.status_bar:
                self.status_bar.showMessage("Versión actualizada.")
            return
        self._update_info = data
        if hasattr(self, "btn_update_hdr"):
            self.btn_update_hdr.setText(f"Actualizar a v{nueva}")
            self.btn_update_hdr.setVisible(True)
        url = data.get("download_url")
        if not url:
            if self.status_bar:
                self.status_bar.showMessage("No se encontró el instalador en la última versión.")
            return
        notas = (data.get("notas") or "").strip()
        msj = (
            f"Nueva versión {nueva} disponible. (Tu versión: {actual})\n\n"
            + (f"{notas}\n\n" if notas else "")
            + "¿Quieres descargarla e instalarla ahora?"
        )
        resp = QMessageBox.question(
            self, "Actualización disponible", msj,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if resp == QMessageBox.StandardButton.Yes:
            self._descargar_actualizacion(url, data.get("asset_name") or "actualizacion.exe")

    def _actualizar_desde_header(self):
        data = getattr(self, "_update_info", None)
        if not data:
            self._comprobar_actualizacion(forzar=True)
            return
        url = data.get("download_url")
        if url:
            self._descargar_actualizacion(url, data.get("asset_name") or "actualizacion.exe")
        else:
            self._comprobar_actualizacion(forzar=True)

    def _on_check_fail(self, mensaje):
        self._update_checking = False
        try:
            self.thread_check.quit()
        except Exception:
            pass
        if self.status_bar:
            self.status_bar.showMessage("No se pudo comprobar actualizaciones.")

    def _descargar_actualizacion(self, url, nombre):
        tmpdir = tempfile.mkdtemp(prefix="ga_update_")
        destino = os.path.join(tmpdir, nombre)
        self._dl_progress = QProgressDialog(
            "Descargando actualización…", None, 0, 100, self
        )
        self._dl_progress.setWindowTitle("Actualización")
        self._dl_progress.setWindowModality(Qt.WindowModality.WindowModal)
        self._dl_progress.setValue(0)
        self._dl_progress.setAutoClose(False)
        self._dl_progress.setAutoReset(False)

        self.thread_dl = QThread(self)
        self.worker_dl = DescargaWorker(url, destino)
        self.worker_dl.moveToThread(self.thread_dl)
        self.thread_dl.started.connect(self.worker_dl.run)
        self.thread_dl.finished.connect(self.thread_dl.deleteLater)
        self.worker_dl.progreso.connect(self._on_dl_progreso)
        self.worker_dl.listo.connect(self._on_dl_ok)
        self.worker_dl.fallo.connect(self._on_dl_fail)
        self.thread_dl.start()

    def _on_dl_progreso(self, done, total):
        if total and done is not None:
            self._dl_progress.setValue(min(int(done * 100 / total), 100))

    def _on_dl_ok(self, ruta):
        try:
            self._dl_progress.setValue(100)
            self._dl_progress.close()
        except Exception:
            pass
        if not _es_frozen():
            QMessageBox.information(
                self, "Actualización descargada",
                f"El archivo se descargó en:\n{ruta}\n\n"
                "Como la app se está ejecutando desde Python, instálala "
                "manualmente (sustituye el .exe o ejecuta el instalador)."
            )
            return
        self._aplicar_actualizacion(ruta)

    def _on_dl_fail(self, mensaje):
        try:
            self._dl_progress.close()
        except Exception:
            pass
        QMessageBox.critical(
            self, "Error de descarga",
            f"No se pudo descargar la actualización:\n{mensaje}"
        )
        if self.status_bar:
            self.status_bar.showMessage("Error al descargar la actualización.")

    def _aplicar_actualizacion(self, ruta):
        try:
            exe = os.path.realpath(sys.executable)
            directorio = os.path.dirname(exe)
            nombre = os.path.basename(exe)
            nuevo = os.path.join(directorio, "_nueva_" + nombre)
            shutil.move(ruta, nuevo)
            bat = os.path.join(directorio, "_actualizar.bat")
            contenido = (
                "@echo off\r\n"
                f'cd /d "{directorio}"\r\n'
                ":espera\r\n"
                f'tasklist /fi "IMAGENAME eq {nombre}" 2>nul | find /i "{nombre}" >nul\r\n'
                "if %errorlevel%==0 ( timeout /t 1 /nobreak >nul & goto espera )\r\n"
                f'move /y "{nuevo}" "{exe}"\r\n'
                f'start "" "{exe}"\r\n'
                'del "%~f0"\r\n'
            )
            with open(bat, "w") as f:
                f.write(contenido)
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(["cmd", "/c", bat], shell=False, creationflags=flags)
            QMessageBox.information(
                self, "Actualización",
                "La actualización se instalará automáticamente al salir.\n"
                "La aplicación se cerrará y se abrirá la nueva versión."
            )
            self.close()
        except Exception as ex:
            QMessageBox.critical(
                self, "Error",
                f"No se pudo preparar la actualización:\n{ex}"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    app.setWindowIcon(QIcon(resolver_ruta("icono.ico")))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())