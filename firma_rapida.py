#!/usr/bin/env python3
"""
Firma Rapida v2.1.1
=================
Quick PDF Digital Signing Tool for Windows.

Signs PDFs (and Word documents!) with your .p12/.pfx certificate
without re-entering the certificate path or password each time.

Features:
  - Remembers your certificate and password (encrypted locally)
  - Open .docx files directly: auto-converts to PDF, then sign
  - Visual PDF preview with draggable signature rectangles
  - Add signatures with a button, delete with Supr/Delete key
  - Multiple signature placements on the same document
  - Replicate a signature to all pages with one click
  - Batch signing: queue multiple files, sign them all at once
  - Page navigation and zoom
  - Configurable signature size and metadata

Requirements:
  pip install pyhanko PyMuPDF Pillow cryptography docx2pdf

Usage:
  python firma_rapida.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys
import tempfile
import subprocess
import shutil
from pathlib import Path
from datetime import datetime
from io import BytesIO

# ── Dependency checks ───────────────────────────────────────────────
_missing = []
try:
    import fitz  # PyMuPDF
except ImportError:
    _missing.append("PyMuPDF")

try:
    from PIL import Image, ImageTk
except ImportError:
    _missing.append("Pillow")

try:
    from cryptography.fernet import Fernet
except ImportError:
    _missing.append("cryptography")

try:
    from pyhanko.sign import signers
    from pyhanko.sign.fields import SigFieldSpec, append_signature_field
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
except ImportError:
    _missing.append("pyhanko")

# Optional: QR stamp style for signature appearance
_has_qr_stamp = False
try:
    from pyhanko.stamp import QRStampStyle

    _has_qr_stamp = True
except ImportError:
    pass

# Optional: Trusted timestamp support
_has_timestamper = False
try:
    from pyhanko.sign.timestamps import HTTPTimeStamper

    _has_timestamper = True
except ImportError:
    pass

# docx2pdf is optional but recommended for Word support
_has_docx2pdf = False
try:
    from docx2pdf import convert as docx2pdf_convert

    _has_docx2pdf = True
except ImportError:
    pass

if _missing:
    print("=" * 60)
    print("  Faltan dependencias. Ejecute:")
    print(f"  pip install {' '.join(_missing)}")
    print("=" * 60)
    try:
        import tkinter as _tk

        _root = _tk.Tk()
        _root.withdraw()
        messagebox.showerror(
            "Dependencias faltantes",
            f"Instale las dependencias ejecutando:\n\n"
            f"pip install {' '.join(_missing)}\n\n"
            f"O ejecute instalar.bat",
        )
    except:
        pass
    sys.exit(1)


# ── Word to PDF Conversion ─────────────────────────────────────────


def _convert_via_word_com_direct(docx_path, pdf_output_path):
    """
    Convert using Word COM automation via dynamic dispatch.
    dynamic.Dispatch skips the gencache step that fails in frozen executables.
    """
    import pythoncom
    import win32com.client.dynamic as dynamic

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = dynamic.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_output_path), FileFormat=17)  # 17 = wdFormatPDF
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def convert_docx_to_pdf(docx_path, pdf_output_path=None):
    """
    Convert a .docx file to PDF.
    Tries direct Word COM first, then docx2pdf, then LibreOffice.
    Returns the path to the generated PDF.
    """
    docx_path = str(docx_path)

    if pdf_output_path is None:
        pdf_output_path = str(Path(docx_path).with_suffix(".pdf"))

    # Method 1: direct Word COM via dynamic dispatch (works in frozen exes)
    errors = []
    try:
        _convert_via_word_com_direct(docx_path, pdf_output_path)
        if os.path.exists(pdf_output_path):
            return pdf_output_path
    except Exception as e:
        errors.append(f"Word COM directo: {e}")

    # Method 2: docx2pdf (fallback, also uses Word COM but via early binding)
    if _has_docx2pdf:
        try:
            import pythoncom

            pythoncom.CoInitialize()
            try:
                docx2pdf_convert(docx_path, pdf_output_path)
            finally:
                pythoncom.CoUninitialize()
            if os.path.exists(pdf_output_path):
                return pdf_output_path
        except Exception as e:
            errors.append(f"docx2pdf: {e}")

    # Method 3: LibreOffice command line (fallback)
    for lo_path in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        shutil.which("soffice") or "",
    ]:
        if lo_path and os.path.isfile(lo_path):
            try:
                out_dir = str(Path(pdf_output_path).parent)
                subprocess.run(
                    [
                        lo_path,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        out_dir,
                        docx_path,
                    ],
                    check=True,
                    capture_output=True,
                    timeout=60,
                )
                lo_output = os.path.join(out_dir, Path(docx_path).stem + ".pdf")
                if os.path.exists(lo_output):
                    if lo_output != pdf_output_path:
                        shutil.move(lo_output, pdf_output_path)
                    return pdf_output_path
            except Exception as e:
                errors.append(f"LibreOffice: {e}")

    error_detail = "\n".join(errors) if errors else "(sin detalles)"
    raise RuntimeError(
        "No se pudo convertir el documento Word a PDF.\n\n"
        f"Errores:\n{error_detail}\n\n"
        "Instale una de estas opciones:\n"
        "  1. pip install docx2pdf  (requiere Microsoft Word)\n"
        "  2. LibreOffice (https://www.libreoffice.org)\n\n"
        "O convierta el .docx a PDF manualmente y abra el PDF."
    )


# ── Configuration Manager ──────────────────────────────────────────

CONFIG_DIR = Path.home() / ".firma_rapida"
CONFIG_FILE = CONFIG_DIR / "config.json"
KEY_FILE = CONFIG_DIR / ".key"


class Config:
    """Manages application settings with encrypted password storage."""

    DEFAULTS = {
        "p12_path": "",
        "password_enc": "",
        "remember_password": True,
        "sig_width": 180,
        "sig_height": 60,
        "sig_reason": "Documento firmado digitalmente",
        "sig_location": "",
        "tsa_url": "http://timestamp.digicert.com",
        "last_dir": "",
    }

    def __init__(self):
        CONFIG_DIR.mkdir(exist_ok=True)
        self._fernet = self._init_fernet()
        self.data = {**self.DEFAULTS}
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, encoding="utf-8") as f:
                    self.data.update(json.load(f))
            except Exception:
                pass

    def _init_fernet(self):
        if KEY_FILE.exists():
            key = KEY_FILE.read_bytes()
        else:
            key = Fernet.generate_key()
            KEY_FILE.write_bytes(key)
        return Fernet(key)

    def save(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def set_password(self, pwd):
        if pwd:
            self.data["password_enc"] = self._fernet.encrypt(pwd.encode()).decode()
        else:
            self.data["password_enc"] = ""

    def get_password(self):
        enc = self.data.get("password_enc", "")
        if not enc:
            return ""
        try:
            return self._fernet.decrypt(enc.encode()).decode()
        except Exception:
            return ""

    def __getitem__(self, key):
        return self.data[key]

    def __setitem__(self, key, val):
        self.data[key] = val


# ── Signature Position ─────────────────────────────────────────────


class SigPos:
    """A signature placement: page index + bounding box in PDF points."""

    def __init__(self, page, x, y, w, h):
        self.page = page  # 0-indexed
        self.x = x  # left edge in PDF coords (points, origin bottom-left)
        self.y = y  # bottom edge in PDF coords
        self.w = w
        self.h = h

    @property
    def box(self):
        """Returns (x1, y1, x2, y2) for pyHanko."""
        return (self.x, self.y, self.x + self.w, self.y + self.h)

    def __str__(self):
        return f"Pag. {self.page + 1}  ({int(self.x)}, {int(self.y)})"


# ── PDF Signing Engine ─────────────────────────────────────────────


def _display_box_to_raw_pdf_box(display_box, rotation, page_rect, mediabox):
    """
    Convert a signature bounding box from GUI display coordinates to raw PDF
    user-space coordinates (as expected by pyHanko's SigFieldSpec).

    The GUI computes positions in the *displayed* page coordinate system, which
    PyMuPDF presents after applying the PDF /Rotate entry.  pyHanko writes the
    field into the *raw* PDF user space (before rotation).  When the two differ
    the field ends up off-page and the signature is invisible.

    display_box : (x1, y1, x2, y2) – y-up, bottom-left origin, display dims.
    rotation    : PDF /Rotate value (0, 90, 180 or 270).
    page_rect   : fitz.Rect – page.rect (always starts at 0,0 in PyMuPDF).
    mediabox    : fitz.Rect – page.mediabox (raw dimensions; may have non-zero origin).
    """
    x1, y1, x2, y2 = display_box
    dw = page_rect.width  # display width  (= raw_h for /Rotate 90 or 270)
    dh = page_rect.height  # display height (= raw_w for /Rotate 90 or 270)
    ox = mediabox.x0  # raw origin offset X (usually 0)
    oy = mediabox.y0  # raw origin offset Y (usually 0)

    rotation = int(rotation) % 360

    if rotation == 0:
        # Display == raw (just add mediabox origin offset).
        rx1, ry1, rx2, ry2 = x1 + ox, y1 + oy, x2 + ox, y2 + oy

    elif rotation == 90:
        # PDF /Rotate 90 = display 90° CW.
        # Forward:  disp_x = raw_y − oy,  disp_y = dh − (raw_x − ox)
        # Inverse:  raw_x  = dh − disp_y + ox,  raw_y = disp_x + oy
        rx1 = dh - y2 + ox
        ry1 = x1 + oy
        rx2 = dh - y1 + ox
        ry2 = x2 + oy

    elif rotation == 180:
        # Forward:  disp_x = dw − (raw_x − ox),  disp_y = dh − (raw_y − oy)
        # Inverse:  raw_x  = dw − disp_x + ox,   raw_y = dh − disp_y + oy
        rx1 = dw - x2 + ox
        ry1 = dh - y2 + oy
        rx2 = dw - x1 + ox
        ry2 = dh - y1 + oy

    elif rotation == 270:
        # PDF /Rotate 270 = display 270° CW (= 90° CCW).
        # Forward:  disp_x = dw − (raw_y − oy),  disp_y = raw_x − ox
        # Inverse:  raw_x  = disp_y + ox,          raw_y = dw − disp_x + oy
        rx1 = y1 + ox
        ry1 = dw - x2 + oy
        rx2 = y2 + ox
        ry2 = dw - x1 + oy

    else:
        rx1, ry1, rx2, ry2 = x1 + ox, y1 + oy, x2 + ox, y2 + oy

    return (rx1, ry1, rx2, ry2)


def _normalize_pdf(input_path):
    """
    Re-save the PDF through PyMuPDF to normalize its internal structure.
    This fixes hybrid xref tables and other quirks that trip up pyHanko.
    Returns the path to a clean temporary PDF (caller must delete it).
    """
    tmp = tempfile.NamedTemporaryFile(
        suffix=".pdf", delete=False, dir=tempfile.gettempdir()
    )
    tmp.close()
    try:
        doc = fitz.open(input_path)
        doc.save(tmp.name, garbage=3, deflate=True)
        doc.close()
        return tmp.name
    except Exception:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
        raise


def sign_pdf(config, input_path, output_path, positions):
    """
    Digitally sign a PDF at the given positions using pyHanko.
    Each position creates a separate digital signature field.
    Tested API pattern:
      1. append_signature_field() to create the field
      2. PdfSigner() with stamp_style + timestamper
      3. pdf_signer.sign_pdf() with appearance_text_params for QR
    """
    p12_path = config["p12_path"]
    pwd = config.get_password()

    signer_obj = signers.SimpleSigner.load_pkcs12(
        pfx_file=p12_path,
        passphrase=pwd.encode("utf-8") if pwd else None,
    )

    # Extract and title-case the signer common name for display (split into two lines)
    _display_name = None
    try:
        cn_raw = signer_obj.signing_cert.subject.native.get("common_name", "")
        if cn_raw:
            words = cn_raw.title().split()
            mid = (len(words) + 1) // 2
            _display_name = " ".join(words[:mid]) + "\n" + " ".join(words[mid:])
    except Exception:
        _display_name = None

    # Set up QR stamp style (signature appearance with QR code)
    stamp_style = None
    if _has_qr_stamp:
        try:
            stamp_style = QRStampStyle(
                stamp_text="%(signer)s\n%(ts)s",
                border_width=1,
            )
        except Exception:
            stamp_style = None

    # Fallback to TextStampStyle if QR is not available
    if stamp_style is None:
        try:
            from pyhanko.stamp import TextStampStyle

            stamp_style = TextStampStyle(
                stamp_text="%(signer)s\n%(ts)s",
                border_width=1,
            )
        except Exception:
            stamp_style = None

    # Set up trusted timestamper (so time doesn't come from local clock)
    timestamper = None
    tsa_url = config.data.get("tsa_url", "").strip()
    if tsa_url and _has_timestamper:
        try:
            timestamper = HTTPTimeStamper(url=tsa_url)
        except Exception:
            timestamper = None

    # Normalize PDF to avoid hybrid xref issues
    normalized = None
    try:
        normalized = _normalize_pdf(input_path)
        current_input = normalized
    except Exception:
        current_input = input_path

    # Read page rotation / rect / mediabox from the (normalized) PDF so we can
    # transform display-space signature boxes to raw PDF user-space coords.
    # Page structure (rotation, mediabox) does not change when appending sigs.
    _page_info = {}  # page_index -> (rotation, page_rect, mediabox)
    try:
        _info_doc = fitz.open(current_input)
        for _pg in range(len(_info_doc)):
            _p = _info_doc[_pg]
            _page_info[_pg] = (_p.rotation, _p.rect, _p.mediabox)
        _info_doc.close()
    except Exception:
        pass

    # Build appearance_text_params for QR stamps
    appearance_params = None
    if _has_qr_stamp and isinstance(stamp_style, QRStampStyle):
        appearance_params = {"url": "https://www.securitydata.net.ec/"}
        if _display_name:
            appearance_params["signer"] = _display_name

    tmp_files = []

    try:
        for i, pos in enumerate(positions):
            is_last = i == len(positions) - 1

            if is_last:
                out_path = output_path
            else:
                tmp = tempfile.NamedTemporaryFile(
                    suffix=".pdf", delete=False, dir=tempfile.gettempdir()
                )
                tmp.close()
                tmp_files.append(tmp.name)
                out_path = tmp.name

            field_name = f"Firma_{i + 1}_{datetime.now().strftime('%H%M%S%f')}"

            # Transform the box from display coords to raw PDF user-space coords.
            # Always apply: handles both page rotation AND non-zero MediaBox origins.
            # CAD-generated PDFs (AutoCAD, MicroStation…) often place the coordinate
            # origin at the centre of the page, e.g. x ∈ [-1192, 1192] instead of
            # [0, 2384].  Without the offset the field lands hundreds of points
            # off-screen and the signature is invisible.
            raw_box = pos.box
            if pos.page in _page_info:
                rot, pg_rect, mb = _page_info[pos.page]
                raw_box = _display_box_to_raw_pdf_box(pos.box, rot, pg_rect, mb)

            with open(current_input, "rb") as inf:
                writer = IncrementalPdfFileWriter(inf, strict=False)

                # Step 1: Create the signature field FIRST
                append_signature_field(
                    writer,
                    sig_field_spec=SigFieldSpec(
                        sig_field_name=field_name,
                        on_page=pos.page,
                        box=raw_box,
                    ),
                )

                # Step 2: Create PdfSigner with style and timestamper
                meta = signers.PdfSignatureMetadata(
                    field_name=field_name,
                    reason=config["sig_reason"] or None,
                    location=config["sig_location"] or None,
                )
                pdf_signer = signers.PdfSigner(
                    meta,
                    signer=signer_obj,
                    stamp_style=stamp_style,
                    timestamper=timestamper,
                )

                # Step 3: Sign (pass appearance_text_params for QR url)
                sign_kwargs = {}
                if appearance_params:
                    sign_kwargs["appearance_text_params"] = appearance_params
                result = pdf_signer.sign_pdf(writer, **sign_kwargs)

                with open(out_path, "wb") as outf:
                    outf.write(result.getbuffer())

            current_input = out_path

    finally:
        for tmp_path in tmp_files:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        if normalized:
            try:
                os.unlink(normalized)
            except OSError:
                pass


# ── Settings Dialog ────────────────────────────────────────────────


class SettingsDialog(tk.Toplevel):
    """Configuration dialog for certificate and signature settings."""

    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("Configuracion - Firma Rapida")
        self.config = config
        self.saved = False
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Certificado (.p12 / .pfx):", font=("", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 3)
        )
        path_frame = ttk.Frame(frame)
        path_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.path_var = tk.StringVar(value=config["p12_path"])
        ttk.Entry(path_frame, textvariable=self.path_var, width=50).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(path_frame, text="Examinar...", command=self._browse_p12).pack(
            side="left", padx=(5, 0)
        )

        ttk.Label(frame, text="Contrasena del certificado:", font=("", 9, "bold")).grid(
            row=2, column=0, sticky="w", pady=(0, 3)
        )
        self.pwd_var = tk.StringVar(value=config.get_password())
        ttk.Entry(frame, textvariable=self.pwd_var, show="*", width=50).grid(
            row=3, column=0, sticky="w", pady=(0, 3)
        )

        self.remember_var = tk.BooleanVar(value=config["remember_password"])
        ttk.Checkbutton(
            frame,
            text="Recordar contrasena (cifrada localmente)",
            variable=self.remember_var,
        ).grid(row=4, column=0, sticky="w", pady=(0, 15))

        ttk.Label(
            frame, text="Tamano de firma (puntos PDF):", font=("", 9, "bold")
        ).grid(row=5, column=0, sticky="w", pady=(0, 3))
        dim_frame = ttk.Frame(frame)
        dim_frame.grid(row=6, column=0, sticky="w", pady=(0, 15))
        ttk.Label(dim_frame, text="Ancho:").pack(side="left")
        self.w_var = tk.IntVar(value=config["sig_width"])
        ttk.Spinbox(dim_frame, from_=50, to=500, textvariable=self.w_var, width=6).pack(
            side="left", padx=(3, 15)
        )
        ttk.Label(dim_frame, text="Alto:").pack(side="left")
        self.h_var = tk.IntVar(value=config["sig_height"])
        ttk.Spinbox(dim_frame, from_=20, to=250, textvariable=self.h_var, width=6).pack(
            side="left", padx=(3, 0)
        )

        ttk.Label(frame, text="Razon de firma:", font=("", 9, "bold")).grid(
            row=7, column=0, sticky="w", pady=(0, 3)
        )
        self.reason_var = tk.StringVar(value=config["sig_reason"])
        ttk.Entry(frame, textvariable=self.reason_var, width=50).grid(
            row=8, column=0, sticky="w", pady=(0, 10)
        )

        ttk.Label(frame, text="Ubicacion:", font=("", 9, "bold")).grid(
            row=9, column=0, sticky="w", pady=(0, 3)
        )
        self.loc_var = tk.StringVar(value=config["sig_location"])
        ttk.Entry(frame, textvariable=self.loc_var, width=50).grid(
            row=10, column=0, sticky="w", pady=(0, 15)
        )

        # TSA URL
        ttk.Label(
            frame, text="Servidor de sellado de tiempo (TSA):", font=("", 9, "bold")
        ).grid(row=11, column=0, sticky="w", pady=(0, 3))
        self.tsa_var = tk.StringVar(
            value=config.data.get("tsa_url", "http://timestamp.digicert.com")
        )
        ttk.Entry(frame, textvariable=self.tsa_var, width=50).grid(
            row=12, column=0, sticky="w", pady=(0, 3)
        )
        ttk.Label(
            frame,
            text="Proporciona hora certificada (no del reloj local). Dejar vacio para desactivar.",
            foreground="#666",
            font=("", 7),
        ).grid(row=13, column=0, sticky="w", pady=(0, 20))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=14, column=0, sticky="e")
        ttk.Button(btn_frame, text="Guardar", command=self._save).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="left")

        self.transient(parent)
        self.wait_window()

    def _browse_p12(self):
        path = filedialog.askopenfilename(
            title="Seleccionar certificado",
            filetypes=[
                ("Certificados PKCS#12", "*.p12 *.pfx"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if path:
            self.path_var.set(path)

    def _save(self):
        p12 = self.path_var.get().strip()
        if p12 and not os.path.isfile(p12):
            messagebox.showwarning("Archivo no encontrado", f"No se encuentra:\n{p12}")
            return

        self.config["p12_path"] = p12
        self.config["remember_password"] = self.remember_var.get()
        if self.remember_var.get():
            self.config.set_password(self.pwd_var.get())
        else:
            self.config.set_password("")
        self.config["sig_width"] = self.w_var.get()
        self.config["sig_height"] = self.h_var.get()
        self.config["sig_reason"] = self.reason_var.get().strip()
        self.config["sig_location"] = self.loc_var.get().strip()
        self.config["tsa_url"] = self.tsa_var.get().strip()
        self.config.save()
        self.saved = True
        self.destroy()


# ── Password Prompt ────────────────────────────────────────────────


class PasswordDialog(tk.Toplevel):
    """Quick password prompt when password isn't saved in config."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Contrasena del certificado")
        self.password = None
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=20)
        frame.pack()

        ttk.Label(frame, text="Ingrese la contrasena de su certificado:").pack(
            anchor="w", pady=(0, 5)
        )
        self.pwd_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.pwd_var, show="*", width=40)
        entry.pack(pady=(0, 10))
        entry.focus_set()
        entry.bind("<Return>", lambda e: self._ok())

        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="Aceptar", command=self._ok).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="left")

        self.transient(parent)
        self.wait_window()

    def _ok(self):
        self.password = self.pwd_var.get()
        self.destroy()


# ── Main Application ───────────────────────────────────────────────


class FirmaRapidaApp:
    """Main GUI application for quick PDF signing."""

    CANVAS_OFFSET = 10  # padding around the page image

    def __init__(self, root):
        self.root = root
        self.root.title("Firma Rapida")
        self.root.geometry("950x720")
        self.root.minsize(750, 550)

        self.config = Config()
        self.pdf_doc = None
        self.pdf_path = None
        self.original_path = None
        self.is_from_word = False
        self._temp_pdf = None
        self.current_page = 0
        self.total_pages = 0
        self.positions = []  # list of SigPos
        self.batch_files = []  # original file paths queued for batch signing
        self.zoom = 1.0
        self.photo = None
        self.page_rect = None

        # Interaction state
        self._placing_mode = False  # True when user clicked "+ Agregar Firma"
        self._dragging_idx = None  # index of signature being dragged, or None
        self._drag_offset_x = 0  # offset from mouse to sig rect origin
        self._drag_offset_y = 0
        self._selected_idx = None  # currently selected signature index

        self._build_ui()
        self._check_first_run()

    def _build_ui(self):
        # ── Toolbar ──────────────────────────
        toolbar = ttk.Frame(self.root, padding=(5, 3))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="Abrir PDF/Word", command=self.open_file).pack(
            side="left", padx=2
        )

        ttk.Separator(toolbar, orient="vertical").pack(
            side="left", fill="y", padx=8, pady=2
        )

        ttk.Button(toolbar, text="<", command=self.prev_page, width=3).pack(side="left")
        self.page_label = ttk.Label(
            toolbar, text="Sin documento", width=18, anchor="center"
        )
        self.page_label.pack(side="left", padx=3)
        ttk.Button(toolbar, text=">", command=self.next_page, width=3).pack(side="left")

        ttk.Separator(toolbar, orient="vertical").pack(
            side="left", fill="y", padx=8, pady=2
        )

        ttk.Label(toolbar, text="Zoom:").pack(side="left")
        self.zoom_var = tk.StringVar(value="Ajustar")
        zoom_cb = ttk.Combobox(
            toolbar,
            textvariable=self.zoom_var,
            values=["Ajustar", "50%", "75%", "100%", "125%", "150%", "200%"],
            width=8,
            state="readonly",
        )
        zoom_cb.pack(side="left", padx=2)
        zoom_cb.bind("<<ComboboxSelected>>", lambda e: self.render_page())

        # Right side of toolbar
        ttk.Button(toolbar, text="Configuracion", command=self.open_settings).pack(
            side="right", padx=2
        )
        self.sign_btn = ttk.Button(toolbar, text="FIRMAR", command=self.do_sign)
        self.sign_btn.pack(side="right", padx=2)

        # ── Main area (PanedWindow) ──────────
        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=(3, 0))

        # Left: PDF canvas with scrollbars
        canvas_frame = ttk.Frame(paned)
        paned.add(canvas_frame, weight=4)

        self.canvas = tk.Canvas(canvas_frame, bg="#d0d0d0", cursor="arrow")
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(
            canvas_frame, orient="horizontal", command=self.canvas.xview
        )
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        # Canvas mouse bindings for select/drag/place
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        # Right: document queue + signature placement
        right_panel = ttk.Frame(paned, padding=5)
        paned.add(right_panel, weight=1)

        # Files to sign
        ttk.Label(right_panel, text="Documentos a firmar:", font=("", 10, "bold")).pack(
            anchor="w", pady=(0, 3)
        )

        files_frame = ttk.Frame(right_panel)
        files_frame.pack(fill="x", pady=(0, 3))

        self.files_listbox = tk.Listbox(
            files_frame, height=4, font=("Consolas", 8), activestyle="dotbox"
        )
        files_vsb = ttk.Scrollbar(
            files_frame, orient="vertical", command=self.files_listbox.yview
        )
        self.files_listbox.configure(yscrollcommand=files_vsb.set)
        self.files_listbox.pack(side="left", fill="both", expand=True)
        files_vsb.pack(side="right", fill="y")

        self.files_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        files_btn_row = ttk.Frame(right_panel)
        files_btn_row.pack(fill="x", pady=(0, 5))
        ttk.Button(
            files_btn_row, text="+ Agregar archivos", command=self.add_batch_files
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(files_btn_row, text="Quitar", command=self.remove_batch_file).pack(
            side="left", padx=(3, 0)
        )

        ttk.Separator(right_panel, orient="horizontal").pack(fill="x", pady=(0, 8))

        # Signature placement
        ttk.Label(right_panel, text="Posicion de firma:", font=("", 10, "bold")).pack(
            anchor="w", pady=(0, 5)
        )

        # "+ Agregar Firma" button
        self.add_btn = ttk.Button(
            right_panel, text="+ Agregar Firma", command=self._enter_placing_mode
        )
        self.add_btn.pack(fill="x", pady=(0, 5))

        list_frame = ttk.Frame(right_panel)
        list_frame.pack(fill="both", expand=True)

        self.pos_listbox = tk.Listbox(list_frame, height=7, font=("Consolas", 9))
        list_vsb = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.pos_listbox.yview
        )
        self.pos_listbox.configure(yscrollcommand=list_vsb.set)
        self.pos_listbox.pack(side="left", fill="both", expand=True)
        list_vsb.pack(side="right", fill="y")

        self.pos_listbox.bind("<<ListboxSelect>>", self._on_pos_select)

        # Delete key on listbox removes selected signature
        self.pos_listbox.bind("<Delete>", lambda e: self.remove_position())
        self.pos_listbox.bind("<BackSpace>", lambda e: self.remove_position())

        btn_row = ttk.Frame(right_panel)
        btn_row.pack(fill="x", pady=(5, 0))
        ttk.Button(
            btn_row, text="Eliminar seleccionada", command=self.remove_position
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(btn_row, text="Limpiar todo", command=self.clear_positions).pack(
            side="left", fill="x", expand=True, padx=(3, 0)
        )

        ttk.Button(
            right_panel,
            text="Replicar en todas las paginas",
            command=self.replicate_to_all_pages,
        ).pack(fill="x", pady=(3, 0))

        # Instructions
        ttk.Separator(right_panel, orient="horizontal").pack(fill="x", pady=10)
        instructions = (
            "Instrucciones:\n"
            "1. Abrir un doc. de referencia\n"
            "2. '+ Agregar archivos' para\n"
            "   anadir mas al lote\n"
            "3. Clic '+ Agregar Firma'\n"
            "   en el documento\n"
            "4. Arrastre para reposicionar\n"
            "5. 'Replicar' para todas\n"
            "   las paginas\n"
            "6. Presione FIRMAR"
        )
        ttk.Label(
            right_panel, text=instructions, foreground="#555", justify="left"
        ).pack(anchor="w")

        # ── Status bar ───────────────────────
        self.status = ttk.Label(
            self.root,
            text="Listo. Abra un PDF o documento Word para comenzar.",
            relief="sunken",
            padding=(5, 3),
        )
        self.status.pack(fill="x", side="bottom")

        # Global Delete key binding
        self.root.bind("<Delete>", lambda e: self.remove_position())

    def _check_first_run(self):
        if not self.config["p12_path"]:
            self.status.config(
                text="Configure su certificado .p12 en Configuracion antes de firmar."
            )

    # ── Placing mode ────────────────────────────────────

    def _enter_placing_mode(self):
        """Activate placement mode: next canvas click adds a signature."""
        if not self.pdf_doc:
            messagebox.showinfo("Sin documento", "Abra un PDF o Word primero.")
            return
        self._placing_mode = True
        self.canvas.config(cursor="crosshair")
        self.add_btn.config(text="Haga clic en el documento...")
        self.status.config(
            text="Modo colocacion: haga clic donde desea la firma. Esc para cancelar."
        )
        self.root.bind("<Escape>", lambda e: self._exit_placing_mode())

    def _exit_placing_mode(self):
        """Exit placement mode without placing."""
        self._placing_mode = False
        self.canvas.config(cursor="arrow")
        self.add_btn.config(text="+ Agregar Firma")
        self.status.config(text="Modo colocacion cancelado.")
        self.root.unbind("<Escape>")

    # ── Canvas mouse handlers ───────────────────────────

    def _find_sig_at(self, cx, cy):
        """Find which signature index is at canvas coords (cx, cy), or None."""
        if not self.page_rect:
            return None
        off = self.CANVAS_OFFSET
        for i, pos in enumerate(self.positions):
            if pos.page != self.current_page:
                continue
            # Convert sig PDF coords to canvas coords
            sx1 = pos.x * self.zoom + off
            sy1 = (self.page_rect.height - pos.y - pos.h) * self.zoom + off
            sx2 = (pos.x + pos.w) * self.zoom + off
            sy2 = (self.page_rect.height - pos.y) * self.zoom + off
            if sx1 <= cx <= sx2 and sy1 <= cy <= sy2:
                return i
        return None

    def _on_press(self, event):
        """Handle mouse press: place (if in placing mode), or start drag."""
        if not self.pdf_doc:
            return

        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)

        # If in placing mode, add a new signature
        if self._placing_mode:
            self._place_signature(cx, cy)
            self._exit_placing_mode()
            return

        # Otherwise, check if clicking on an existing signature to select/drag
        idx = self._find_sig_at(cx, cy)
        if idx is not None:
            self._selected_idx = idx
            self._dragging_idx = idx
            # Highlight in listbox
            self.pos_listbox.selection_clear(0, "end")
            self.pos_listbox.selection_set(idx)
            self.pos_listbox.see(idx)
            # Calculate drag offset (distance from mouse to sig top-left corner)
            pos = self.positions[idx]
            off = self.CANVAS_OFFSET
            sx1 = pos.x * self.zoom + off
            sy1 = (self.page_rect.height - pos.y - pos.h) * self.zoom + off
            self._drag_offset_x = cx - sx1
            self._drag_offset_y = cy - sy1
            self.canvas.config(cursor="fleur")  # move cursor
            self.status.config(text=f"Arrastrando firma: {pos}")
        else:
            # Clicked on empty area: deselect
            self._selected_idx = None
            self._dragging_idx = None
            self.pos_listbox.selection_clear(0, "end")

    def _on_drag(self, event):
        """Handle mouse drag: move the signature rectangle."""
        if self._dragging_idx is None:
            return

        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        off = self.CANVAS_OFFSET

        pos = self.positions[self._dragging_idx]

        # New top-left in canvas coords
        new_sx1 = cx - self._drag_offset_x
        new_sy1 = cy - self._drag_offset_y

        # Convert canvas coords back to PDF coords
        new_pdf_x = (new_sx1 - off) / self.zoom
        new_pdf_y_top = (new_sy1 - off) / self.zoom
        new_pdf_y = self.page_rect.height - new_pdf_y_top - pos.h

        # Clamp to page bounds
        new_pdf_x = max(0, min(new_pdf_x, self.page_rect.width - pos.w))
        new_pdf_y = max(0, min(new_pdf_y, self.page_rect.height - pos.h))

        pos.x = new_pdf_x
        pos.y = new_pdf_y

        # Update listbox text
        self.pos_listbox.delete(self._dragging_idx)
        self.pos_listbox.insert(self._dragging_idx, str(pos))
        self.pos_listbox.selection_set(self._dragging_idx)

        # Redraw overlays only (not the whole page image for performance)
        self.canvas.delete("sig")
        for p in self.positions:
            if p.page == self.current_page:
                self._draw_sig_rect(p, highlight=(p is pos))

    def _on_release(self, event):
        """Handle mouse release: finish drag."""
        if self._dragging_idx is not None:
            pos = self.positions[self._dragging_idx]
            self.status.config(text=f"Firma reposicionada: {pos}")
            self._dragging_idx = None
            self.canvas.config(cursor="arrow")

    def _place_signature(self, cx, cy):
        """Place a new signature at canvas coordinates."""
        off = self.CANVAS_OFFSET
        sig_w = self.config["sig_width"]
        sig_h = self.config["sig_height"]

        # Convert canvas coords to PDF coords (center the sig on the click)
        pdf_x = (cx - off) / self.zoom - sig_w / 2
        py_from_top = (cy - off) / self.zoom
        pdf_y = self.page_rect.height - py_from_top - sig_h / 2

        # Clamp to page boundaries
        pdf_x = max(0, min(pdf_x, self.page_rect.width - sig_w))
        pdf_y = max(0, min(pdf_y, self.page_rect.height - sig_h))

        pos = SigPos(self.current_page, pdf_x, pdf_y, sig_w, sig_h)
        self.positions.append(pos)
        self.pos_listbox.insert("end", str(pos))
        self._draw_sig_rect(pos)
        self._selected_idx = len(self.positions) - 1
        self.pos_listbox.selection_clear(0, "end")
        self.pos_listbox.selection_set(self._selected_idx)
        self.status.config(text=f"Firma anadida: {pos}")

    # ── Drawing ─────────────────────────────────────────

    def _draw_sig_rect(self, pos, highlight=False):
        """Draw a signature rectangle overlay on the canvas."""
        off = self.CANVAS_OFFSET
        x1 = pos.x * self.zoom + off
        y1 = (self.page_rect.height - pos.y - pos.h) * self.zoom + off
        x2 = (pos.x + pos.w) * self.zoom + off
        y2 = (self.page_rect.height - pos.y) * self.zoom + off

        color = "#CC3300" if highlight else "#0055AA"
        self.canvas.create_rectangle(
            x1,
            y1,
            x2,
            y2,
            outline=color,
            width=2,
            dash=(5, 3),
            tags="sig",
        )
        cx_r, cy_r = (x1 + x2) / 2, (y1 + y2) / 2
        self.canvas.create_text(
            cx_r,
            cy_r,
            text="Firma",
            fill=color,
            font=("", 8, "bold"),
            tags="sig",
        )

    # ── File operations ──────────────────────────────────

    def _cleanup_temp_pdf(self):
        if self._temp_pdf and os.path.exists(self._temp_pdf):
            try:
                os.unlink(self._temp_pdf)
            except OSError:
                pass
            self._temp_pdf = None

    def open_file(self):
        """Open one document and start a new batch session."""
        init_dir = self.config["last_dir"] or str(Path.home())
        path = filedialog.askopenfilename(
            initialdir=init_dir,
            title="Abrir documento PDF o Word",
            filetypes=[
                ("PDF y Word", "*.pdf *.docx *.doc"),
                ("Archivos PDF", "*.pdf"),
                ("Documentos Word", "*.docx *.doc"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not path:
            return

        self.config["last_dir"] = str(Path(path).parent)
        self.config.save()

        # New session: clear batch list and signature positions
        self.batch_files.clear()
        self.files_listbox.delete(0, "end")
        self.positions.clear()
        self.pos_listbox.delete(0, "end")
        self._selected_idx = None

        # Load the file for preview (handles Word conversion internally)
        self._load_for_preview(path)
        if not self.pdf_doc:
            return  # loading failed

        # Register in batch list
        self.batch_files.append(path)
        self.files_listbox.insert("end", Path(path).name)
        self.files_listbox.selection_set(0)

        # Auto-enter placing mode
        self._enter_placing_mode()
        name = Path(path).name
        word_note = " (convertido de Word)" if self.is_from_word else ""
        self.status.config(
            text=f"Abierto: {name}{word_note} ({self.total_pages} pag.) - Haga clic para colocar la primera firma"
        )

    def _load_for_preview(self, path):
        """Load a file into the canvas preview without resetting the batch list or positions."""
        if self.pdf_doc:
            self.pdf_doc.close()
            self.pdf_doc = None
        self._cleanup_temp_pdf()

        self.original_path = path
        self.is_from_word = False

        ext = Path(path).suffix.lower()
        if ext in (".docx", ".doc"):
            self.is_from_word = True
            self.status.config(text=f"Convirtiendo {Path(path).name} a PDF...")
            self.root.update_idletasks()
            try:
                tmp_pdf = tempfile.NamedTemporaryFile(
                    suffix=".pdf",
                    delete=False,
                    dir=str(Path(path).parent),
                    prefix=f"{Path(path).stem}_tmp_",
                )
                tmp_pdf.close()
                self._temp_pdf = tmp_pdf.name
                convert_docx_to_pdf(path, self._temp_pdf)
                pdf_to_open = self._temp_pdf
            except Exception as e:
                messagebox.showerror(
                    "Error al convertir Word",
                    f"No se pudo convertir el documento Word a PDF:\n\n{e}",
                )
                self._cleanup_temp_pdf()
                return
        else:
            pdf_to_open = path

        try:
            self.pdf_doc = fitz.open(pdf_to_open)
        except Exception as e:
            messagebox.showerror("Error al abrir PDF", str(e))
            self._cleanup_temp_pdf()
            return

        self.pdf_path = pdf_to_open
        self.current_page = 0
        self.total_pages = len(self.pdf_doc)
        self.render_page()
        name = Path(path).name
        word_note = " (Word)" if self.is_from_word else ""
        self.root.title(f"Firma Rapida - {name}")
        if path in self.batch_files:
            idx = self.batch_files.index(path)
            self.files_listbox.selection_clear(0, "end")
            self.files_listbox.selection_set(idx)
        self.status.config(
            text=f"Previsualizando: {name}{word_note} ({self.total_pages} pag.)"
        )

    def add_batch_files(self):
        """Add more documents to the batch signing queue."""
        init_dir = self.config["last_dir"] or str(Path.home())
        files = filedialog.askopenfilenames(
            initialdir=init_dir,
            title="Agregar documentos al lote",
            filetypes=[
                ("PDF y Word", "*.pdf *.docx *.doc"),
                ("Archivos PDF", "*.pdf"),
                ("Documentos Word", "*.docx *.doc"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not files:
            return
        self.config["last_dir"] = str(Path(files[0]).parent)
        self.config.save()
        added = 0
        for f in files:
            if f not in self.batch_files:
                self.batch_files.append(f)
                self.files_listbox.insert("end", Path(f).name)
                added += 1
        # If no file is previewed yet, load the first one
        if not self.pdf_doc and self.batch_files:
            self._load_for_preview(self.batch_files[0])
        self.status.config(
            text=f"{added} archivo(s) agregado(s). Total en lote: {len(self.batch_files)}."
        )

    def remove_batch_file(self):
        """Remove the selected file from the batch queue."""
        sel = self.files_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        removed = self.batch_files.pop(idx)
        self.files_listbox.delete(idx)
        # If the removed file was being previewed, switch to another or clear canvas
        if removed == self.original_path:
            if self.batch_files:
                self._load_for_preview(self.batch_files[0])
            else:
                if self.pdf_doc:
                    self.pdf_doc.close()
                    self.pdf_doc = None
                self._cleanup_temp_pdf()
                self.pdf_path = None
                self.original_path = None
                self.canvas.delete("all")
                self.page_label.config(text="Sin documento")
                self.root.title("Firma Rapida")
        n = len(self.batch_files)
        self.status.config(text=f"Archivo quitado. Quedan {n} en el lote.")

    def _on_file_select(self, event):
        """Preview the file selected in the documents listbox."""
        sel = self.files_listbox.curselection()
        if not sel:
            return
        path = self.batch_files[sel[0]]
        if path != self.original_path:
            self._load_for_preview(path)

    def render_page(self):
        if not self.pdf_doc:
            return

        page = self.pdf_doc[self.current_page]
        self.page_rect = page.rect

        zoom_text = self.zoom_var.get()
        if zoom_text == "Ajustar":
            canvas_w = max(self.canvas.winfo_width(), 400)
            canvas_h = max(self.canvas.winfo_height(), 500)
            zx = (canvas_w - 30) / self.page_rect.width
            zy = (canvas_h - 30) / self.page_rect.height
            self.zoom = min(zx, zy, 2.5)
        else:
            self.zoom = int(zoom_text.replace("%", "")) / 100.0

        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.photo = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        off = self.CANVAS_OFFSET
        self.canvas.create_image(off, off, anchor="nw", image=self.photo, tags="page")
        self.canvas.configure(
            scrollregion=(0, 0, pix.width + 2 * off, pix.height + 2 * off)
        )

        for pos in self.positions:
            if pos.page == self.current_page:
                self._draw_sig_rect(pos)

        self.page_label.config(
            text=f"Pag. {self.current_page + 1} / {self.total_pages}"
        )

    # ── Listbox / selection interactions ─────────────────

    def _on_pos_select(self, event):
        """Jump to the page of the selected signature position."""
        sel = self.pos_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self._selected_idx = idx
        if idx < len(self.positions):
            target_page = self.positions[idx].page
            if target_page != self.current_page:
                self.current_page = target_page
                self.render_page()
            else:
                # Redraw to highlight selection
                self.canvas.delete("sig")
                for i, p in enumerate(self.positions):
                    if p.page == self.current_page:
                        self._draw_sig_rect(p, highlight=(i == idx))

    def remove_position(self):
        """Remove the currently selected signature."""
        sel = self.pos_listbox.curselection()
        if not sel:
            # Try using _selected_idx
            if self._selected_idx is not None and self._selected_idx < len(
                self.positions
            ):
                idx = self._selected_idx
            else:
                return
        else:
            idx = sel[0]

        self.positions.pop(idx)
        self.pos_listbox.delete(idx)
        self._selected_idx = None
        self.render_page()
        self.status.config(text="Firma eliminada")

    def clear_positions(self):
        self.positions.clear()
        self.pos_listbox.delete(0, "end")
        self._selected_idx = None
        if self.pdf_doc:
            self.render_page()
        self.status.config(text="Todas las firmas eliminadas")

    def replicate_to_all_pages(self):
        """Replicate the current page's signatures to every page of the document."""
        if not self.pdf_doc:
            messagebox.showinfo("Sin documento", "Abra un PDF primero.")
            return
        current_sigs = [p for p in self.positions if p.page == self.current_page]
        if not current_sigs:
            messagebox.showinfo(
                "Sin firmas",
                f"No hay firmas en la pagina {self.current_page + 1} para replicar.",
            )
            return
        if self.total_pages == 1:
            messagebox.showinfo(
                "Una sola pagina", "El documento solo tiene una pagina."
            )
            return
        if not messagebox.askyesno(
            "Replicar firmas",
            f"Replicar {len(current_sigs)} firma(s) de la pagina "
            f"{self.current_page + 1} a las {self.total_pages} paginas del documento?\n\n"
            "Las firmas existentes en otras paginas seran reemplazadas.",
        ):
            return
        new_positions = []
        for page_idx in range(self.total_pages):
            for sig in current_sigs:
                new_positions.append(SigPos(page_idx, sig.x, sig.y, sig.w, sig.h))
        self.positions = new_positions
        self.pos_listbox.delete(0, "end")
        for p in self.positions:
            self.pos_listbox.insert("end", str(p))
        self._selected_idx = None
        self.render_page()
        self.status.config(
            text=f"Firma replicada en {self.total_pages} paginas "
            f"({len(self.positions)} posiciones en total)."
        )

    def prev_page(self):
        if self.pdf_doc and self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def next_page(self):
        if self.pdf_doc and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.render_page()

    def open_settings(self):
        SettingsDialog(self.root, self.config)

    # ── Signing ─────────────────────────────────────────

    def do_sign(self):
        """Sign all documents in the batch list at the placed signature positions."""
        if not self.batch_files:
            messagebox.showwarning("Sin documentos", "Agregue al menos un documento.")
            return
        if not self.positions:
            messagebox.showwarning(
                "Sin firmas",
                "Use '+ Agregar Firma' para colocar al menos una firma.",
            )
            return
        if not self.config["p12_path"]:
            messagebox.showwarning(
                "Sin certificado",
                "Configure su certificado .p12 en Configuracion.",
            )
            self.open_settings()
            return
        if not os.path.isfile(self.config["p12_path"]):
            messagebox.showerror(
                "Certificado no encontrado",
                f"No se encuentra el archivo:\n{self.config['p12_path']}",
            )
            return

        pwd = self.config.get_password()
        if not pwd:
            dlg = PasswordDialog(self.root)
            if dlg.password is None:
                return
            pwd = dlg.password
            self.config.set_password(pwd)

        if self.pdf_doc:
            self.pdf_doc.close()
            self.pdf_doc = None

        # ── Single file: ask where to save ─────────────────────────────────
        if len(self.batch_files) == 1:
            orig = self.batch_files[0]
            pdf_src = (
                self.pdf_path
                if (self.pdf_path and os.path.exists(str(self.pdf_path)))
                else orig
            )
            stem = Path(orig).stem
            default_name = f"{stem}_firmado.pdf"
            output = filedialog.asksaveasfilename(
                initialdir=str(Path(orig).parent),
                initialfile=default_name,
                title="Guardar PDF firmado como...",
                filetypes=[("PDF", "*.pdf")],
                defaultextension=".pdf",
            )
            if not output:
                self._load_for_preview(orig)
                return

            _first_sig_page = self.positions[0].page if self.positions else 0
            _signed_output = None
            try:
                self.status.config(text="Firmando documento... por favor espere.")
                self.root.update_idletasks()
                sign_pdf(self.config, pdf_src, output, self.positions)
                _signed_output = output

                # ── Diagnostic: verify signature field placement ──
                _diag_warn = ""
                try:
                    _d = fitz.open(output)
                    for _pgi in range(len(_d)):
                        _pg = _d[_pgi]
                        _pr = _pg.rect
                        for _w in _pg.widgets():
                            _wr = _w.rect
                            if (
                                _wr.x1 <= _pr.x0
                                or _wr.x0 >= _pr.x1
                                or _wr.y1 <= _pr.y0
                                or _wr.y0 >= _pr.y1
                            ):
                                _diag_warn += (
                                    f"\n\u26a0 Pag.{_pgi+1}: rect de firma "
                                    f"({_wr.x0:.0f},{_wr.y0:.0f},{_wr.x1:.0f},{_wr.y1:.0f})"
                                    f" fuera de la pagina "
                                    f"({_pr.x0:.0f},{_pr.y0:.0f},{_pr.x1:.0f},{_pr.y1:.0f})"
                                )
                    _d.close()
                except Exception:
                    pass

                msg = f"Documento firmado correctamente.\n\n{output}"
                if _diag_warn:
                    msg += (
                        "\n\n\u26a0 ATENCION: la firma fue colocada fuera del area"
                        " visible de la pagina. Es posible que no se vea en el"
                        " visor PDF." + _diag_warn
                    )
                self.status.config(text=f"Documento firmado: {Path(output).name}")
                messagebox.showinfo("Firma exitosa", msg)
                self.positions.clear()
                self.pos_listbox.delete(0, "end")
                self._selected_idx = None

            except Exception as e:
                messagebox.showerror("Error al firmar", f"Ocurrio un error:\n\n{e}")
                self.status.config(text=f"Error: {e}")

            finally:
                show_path = (
                    _signed_output
                    if (_signed_output and os.path.exists(_signed_output))
                    else pdf_src
                )
                if show_path and os.path.exists(str(show_path)):
                    self.pdf_doc = fitz.open(show_path)
                    self.pdf_path = show_path
                    self.total_pages = len(self.pdf_doc)
                    self.current_page = min(_first_sig_page, self.total_pages - 1)
                    self.render_page()
                    if _signed_output:
                        self.status.config(
                            text="Mostrando PDF firmado. Verifique la firma en el visor."
                        )

        # ── Multiple files: batch sign, auto-save next to each original ────────
        else:
            ok = 0
            batch_errors = []
            for i, orig in enumerate(self.batch_files):
                self.status.config(
                    text=f"Firmando {i + 1}/{len(self.batch_files)}: {Path(orig).name}..."
                )
                self.root.update_idletasks()
                tmp_pdf = None
                try:
                    ext = Path(orig).suffix.lower()
                    if ext in (".docx", ".doc"):
                        t = tempfile.NamedTemporaryFile(
                            suffix=".pdf",
                            delete=False,
                            dir=tempfile.gettempdir(),
                            prefix=f"{Path(orig).stem}_sign_",
                        )
                        t.close()
                        tmp_pdf = t.name
                        convert_docx_to_pdf(orig, tmp_pdf)
                        pdf_src = tmp_pdf
                    else:
                        pdf_src = orig

                    info = fitz.open(pdf_src)
                    n_pages = len(info)
                    info.close()

                    valid_pos = [p for p in self.positions if p.page < n_pages]
                    if not valid_pos:
                        batch_errors.append(
                            f"{Path(orig).name}: ninguna firma aplica "
                            f"(doc tiene {n_pages} pag.)"
                        )
                        continue

                    out = str(Path(orig).parent / f"{Path(orig).stem}_firmado.pdf")
                    sign_pdf(self.config, pdf_src, out, valid_pos)
                    ok += 1

                except Exception as exc:
                    batch_errors.append(f"{Path(orig).name}: {exc}")
                finally:
                    if tmp_pdf and os.path.exists(tmp_pdf):
                        try:
                            os.unlink(tmp_pdf)
                        except OSError:
                            pass

            summary = f"Firmados: {ok} de {len(self.batch_files)} documento(s)."
            if batch_errors:
                summary += "\n\nErrores:\n" + "\n".join(f"• {e}" for e in batch_errors)
                messagebox.showwarning("Firma en lote completada", summary)
            else:
                messagebox.showinfo("Firma en lote completada", summary)
            self.positions.clear()
            self.pos_listbox.delete(0, "end")
            self._selected_idx = None
            self.status.config(
                text=f"Lote completado: {ok}/{len(self.batch_files)} firmados."
            )
            # Reload preview of first file
            if self.batch_files:
                self._load_for_preview(self.batch_files[0])


# ── Entry point ────────────────────────────────────────────────────


def main():
    root = tk.Tk()

    style = ttk.Style()
    for theme in ("vista", "clam", "winnative"):
        if theme in style.theme_names():
            style.theme_use(theme)
            break

    app = FirmaRapidaApp(root)

    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()


if __name__ == "__main__":
    main()
