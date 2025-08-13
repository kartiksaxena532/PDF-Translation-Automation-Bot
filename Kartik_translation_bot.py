# app_gui.py  — Local desktop GUI -> EXE (blazing-fast: no artificial delays)
import sys, os, re, time, subprocess, asyncio, shutil
from urllib.request import urlopen
from urllib.error import URLError
from pathlib import Path
from mimetypes import guess_type
from threading import Thread
from typing import Optional

import img2pdf
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image
from playwright.async_api import async_playwright, TimeoutError as PWTimeout
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional OCR deps (detected at runtime)
try:
    import pytesseract  # fallback OCR engine wrapper
except Exception:
    pytesseract = None

try:
    from PyPDF2 import PdfMerger  # fallback PDF merge for pytesseract path
except Exception:
    PdfMerger = None

# ------------ Configuration ------------
APP_NAME = "PDF Translator Bot - By Kartik Saxena"
RAW_DIR   = Path(r"H:\Bot\raw images")
TRANS_DIR = Path(r"H:\Bot\translated images")
RAW_DIR.mkdir(parents=True, exist_ok=True)
TRANS_DIR.mkdir(parents=True, exist_ok=True)

# POPPLER location resolver: use bundled folder in EXE, else fallback dev path
def resolve_poppler_bin() -> Path:
    if getattr(sys, "_MEIPASS", None):  # running from PyInstaller bundle
        return Path(sys._MEIPASS) / "poppler_bin"
    # ↓ edit this fallback once for dev runs:
    return Path(r"H:\Downloads\Oppo Amos EP040\Release-24.08.0-0\poppler-24.08.0\Library\bin")

# Chrome remote debugging attach
DEBUG_PORT = 9222
REMOTE     = f"http://localhost:{DEBUG_PORT}"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
EDGE_CANDIDATES = [
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]
USER_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", r"C:\Users\Public")) / "Chrome" / "PWProfile"

IMG_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}

# ------------ Helpers ------------
def dbg_ready() -> bool:
    try:
        with urlopen(f"{REMOTE}/json/version", timeout=1.5) as r:
            return r.status == 200
    except URLError:
        return False

def find_browser_exe() -> str:
    for p in CHROME_CANDIDATES + EDGE_CANDIDATES:
        if Path(p).is_file():
            return p
    raise FileNotFoundError("Chrome/Edge not found. Update CHROME_CANDIDATES/EDGE_CANDIDATES.")

def launch_chrome_if_needed(log):
    if dbg_ready():
        log("Chrome debug port is ready.")
        return
    exe = find_browser_exe()
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    args = [
        exe,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={str(USER_DATA_DIR)}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
    log("Launching Chrome ...Grab A Coffee , Sit Back & Relax")
    for _ in range(60):
        if dbg_ready():
            log("Let me do the boring stuff now!Meanwhile you can disturb your co-worker 😈 Just Suggesting 😂")
            return
        time.sleep(0.25)
    raise TimeoutError("Could not start Chrome with remote debugging port.")

def wipe_images_only(folder: Path):
    if not folder.exists():
        return
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try: p.unlink(missing_ok=True)
            except: pass

def which(cmd: str) -> Optional[str]:
    """Portable which()."""
    p = shutil.which(cmd)
    return p if p and Path(p).exists() else None

# ------------ Core pipeline ------------
def extract_pages(pdf_path: str, dpi: int, log) -> list[Path]:
    poppler_bin = resolve_poppler_bin()
    info = pdfinfo_from_path(pdf_path, poppler_path=str(poppler_bin))
    log(f"Processing your PDF which has {info['Pages']} pages → saving PNGs in {RAW_DIR}")
    shutil.rmtree(RAW_DIR, ignore_errors=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    log("Wait for it...")
    out_paths: list[Path] = []
    for i, pg in enumerate(convert_from_path(pdf_path, dpi=dpi, fmt="png", poppler_path=str(poppler_bin)), 1):
        out = RAW_DIR / f"page-{i:03}.png"
        pg.save(out); out_paths.append(out)
        log(f"  ✓ {out.name}")
    return out_paths

async def translate_images(img_paths: list[Path], target_lang: str, close_browser: bool, log=None) -> list[Path]:
    """
    FAST PATH:
    - No artificial sleeps.
    - Only wait for required UI states (buttons/labels) to become available.
    """
    shutil.rmtree(TRANS_DIR, ignore_errors=True)
    TRANS_DIR.mkdir(parents=True, exist_ok=True)
    results: list[Path] = []

    url = f"https://translate.google.co.in/?sl=auto&tl={target_lang}&op=images"
    launch_chrome_if_needed(log or (lambda *_: None))

    async with async_playwright() as pw:
        browser = await pw.chromium.connect_over_cdp(REMOTE)
        try:
            ctx = browser.contexts[0] if browser.contexts else await browser.new_context(accept_downloads=True)
            page = await ctx.new_page()
            await page.goto(url)

            browse_btn   = page.get_by_role("button", name=re.compile(r"Browse your files", re.I))
            download_btn = page.get_by_role("button", name=re.compile(r"(Download translation|Download)", re.I))
            clear_btn    = page.get_by_role("button", name=re.compile(r"(Clear image|Clear)", re.I))
            show_translated = page.get_by_text("Show translated", exact=True)

            total = len(img_paths)
            for idx, img in enumerate(img_paths, 1):
                log(f"[{idx}/{total}] {img.name}")

                payload = {"name": img.name, "mimeType": guess_type(img.name)[0] or "application/octet-stream", "buffer": img.read_bytes()}
                # Prefer direct file input (fast). Fallback to file chooser if needed.
                try:
                    await page.locator('input[type="file"]').set_input_files(payload, timeout=1500)
                except Exception:
                    async with page.expect_file_chooser() as fc:
                        await browse_btn.click()
                    chooser = await fc.value
                    await chooser.set_files(payload)

                # If we see "Show translated", we're still on ORIGINAL → click it immediately
                if await show_translated.count() > 0:
                    await show_translated.first.click()

                # Wait for download to be ready, then download immediately
                await download_btn.first.wait_for(state="visible", timeout=60000)
                async with page.expect_download() as dl_info:
                    await download_btn.first.click()
                dl = await dl_info.value
                suggested = dl.suggested_filename or f"{img.stem}-translated.png"
                ext = Path(suggested).suffix or ".png"
                out_img = TRANS_DIR / f"{img.stem}-translated{ext}"
                await dl.save_as(str(out_img))
                results.append(out_img)
                log(f"   ↳ saved {out_img.name}")

                # Clear for next image (no sleep)
                try: await clear_btn.click()
                except: pass

            await page.close()
        finally:
            if close_browser:
                try:
                    await browser.close()
                    log("Closed Chrome window.")
                except: pass

    return results

def build_pdf(images: list[Path], out_pdf: Path, log):
    log("Building Your Final PDF…")
    normalized: list[Path] = []
    for p in images:
        if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
            normalized.append(p); continue
        with Image.open(p) as im:
            q = p.with_suffix(".png"); im.save(q); normalized.append(q)
    with open(out_pdf, "wb") as f:
        f.write(img2pdf.convert([str(p) for p in normalized]))
    log(f"✓ Saved → {out_pdf}")

# ------------ OCR helpers ------------
def run_ocrmypdf(in_pdf: Path, out_pdf: Path, ocr_lang: str, log) -> bool:
    """
    Try to OCR with ocrmypdf. Returns True if it succeeded.
    """
    exe = which("ocrmypdf")
    if not exe:
        return False

    # ocrmypdf arguments:
    # --skip-text : don't re-OCR pages that already have text
    # --force-ocr : ensure an OCR layer is created on image pages
    # --optimize 3: decent optimization balance
    # --jobs N   : parallelism
    cmd = [
        exe, "--skip-text", "--force-ocr",
        "--optimize", "3",
        "--jobs", "4",
        "--language", ocr_lang or "eng",
        str(in_pdf), str(out_pdf)
    ]
    log("Running OCR with ocrmypdf…")
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        log("✓ OCR complete via ocrmypdf.")
        return True
    except Exception as e:
        log(f"⚠ ocrmypdf failed ({e}). Will try pytesseract fallback if available.")
        return False

def build_pdf_with_pytesseract(images: list[Path], out_pdf: Path, ocr_lang: str, log):
    """
    Fallback OCR: generate per-page searchable PDFs with pytesseract and merge.
    """
    if pytesseract is None or PdfMerger is None:
        raise RuntimeError("pytesseract/PyPDF2 not installed; cannot fallback OCR.")

    # If needed, you can hard-set the Tesseract binary path:
    # pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    tmp_pdfs: list[Path] = []
    log("Building searchable PDF with pytesseract (this may take a while)…")
    try:
        for i, img in enumerate(images, 1):
            log(f"  OCR [{i}/{len(images)}] {img.name}")
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(
                str(img), lang=(ocr_lang or "eng"), extension='pdf'
            )
            p = img.with_suffix(".page.pdf")
            with open(p, "wb") as f:
                f.write(pdf_bytes)
            tmp_pdfs.append(p)

        merger = PdfMerger()
        for p in tmp_pdfs:
            merger.append(str(p))
        merger.write(str(out_pdf))
        merger.close()
        log("✓ OCR complete via pytesseract.")
    finally:
        for p in tmp_pdfs:
            try: p.unlink(missing_ok=True)
            except: pass

async def translate_pdf(
    input_pdf: str,
    output_pdf: str,
    target_lang: str = "en",
    dpi: int = 150,
    close_browser: bool = True,
    log=print,
    use_ocr: bool = True,          # NEW: default to True (make searchable)
    ocr_lang: str = "eng",         # NEW: Tesseract/ocrmypdf language
):
    if not Path(input_pdf).is_file():
        raise FileNotFoundError(f"File not found: {input_pdf}")
    pages = extract_pages(input_pdf, dpi, log)
    translated = await translate_images(pages, target_lang, close_browser, log)

    # Build an image-only PDF first (fast); OCR will make it searchable
    temp_image_pdf = Path(output_pdf).with_suffix(".imageonly.pdf")
    build_pdf(translated, temp_image_pdf, log)

    final_pdf = Path(output_pdf)
    if use_ocr:
        ok = run_ocrmypdf(temp_image_pdf, final_pdf, ocr_lang, log)
        if not ok:
            # Fallback OCR using pytesseract: build final directly from images
            build_pdf_with_pytesseract(translated, final_pdf, ocr_lang, log)
        # Clean intermediate
        try: temp_image_pdf.unlink(missing_ok=True)
        except: pass
    else:
        # No OCR requested → keep image-only as final
        try:
            if final_pdf.exists():
                final_pdf.unlink()
        except: pass
        temp_image_pdf.replace(final_pdf)

    # cleanup
    wipe_images_only(RAW_DIR)
    wipe_images_only(TRANS_DIR)
    log("Cleaned up temporary images.")

# ------------ GUI ------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("900x420")
        self.resizable(False, False)

        self.in_var  = tk.StringVar()
        self.out_var = tk.StringVar(value=str((TRANS_DIR / "translated.pdf")))
        self.lang    = tk.StringVar(value="en")
        self.dpi     = tk.IntVar(value=150)
        self.close_chrome = tk.BooleanVar(value=True)
        self.use_ocr = tk.BooleanVar(value=True)     # NEW
        self.ocr_lang = tk.StringVar(value="eng")    # NEW

        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        # row 1: input
        ttk.Label(frm, text="Input PDF:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.in_var, width=80).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Browse…", command=self.pick_in).grid(row=0, column=2)
        # row 2: output
        ttk.Label(frm, text="Output PDF:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.out_var, width=80).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Choose…", command=self.pick_out).grid(row=1, column=2)
        # row 3: options
        opt = ttk.Frame(frm); opt.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8,4))
        ttk.Label(opt, text="Target language (tl):").grid(row=0, column=0, sticky="w")
        ttk.Entry(opt, textvariable=self.lang, width=8).grid(row=0, column=1, padx=6)
        ttk.Label(opt, text="DPI:").grid(row=0, column=2, sticky="w")
        ttk.Entry(opt, textvariable=self.dpi, width=6).grid(row=0, column=3, padx=6)
        ttk.Checkbutton(opt, text="Close Chrome on finish", variable=self.close_chrome).grid(row=0, column=4, padx=12)

        # NEW OCR controls
        ttk.Checkbutton(opt, text="Make PDF searchable (OCR)", variable=self.use_ocr).grid(row=0, column=5, padx=12)
        ttk.Label(opt, text="OCR lang:").grid(row=0, column=6, sticky="w")
        ttk.Entry(opt, textvariable=self.ocr_lang, width=8).grid(row=0, column=7, padx=6)

        # row 4: buttons
        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=3, pady=6, sticky="ew")
        self.run_btn = ttk.Button(btns, text="Translate PDF", command=self.start)
        self.run_btn.pack(side="left")
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="right")

        # row 5: log
        self.log = tk.Text(frm, height=14); self.log.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(6,0))
        frm.columnconfigure(1, weight=1)

    def pick_in(self):
        f = filedialog.askopenfilename(filetypes=[("PDF files","*.pdf")])
        if f: self.in_var.set(f)

    def pick_out(self):
        f = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files","*.pdf")])
        if f: self.out_var.set(f)

    def log_write(self, msg):
        self.log.insert("end", msg + "\n"); self.log.see("end"); self.update_idletasks()

    def start(self):
        if not self.in_var.get():
            messagebox.showwarning(APP_NAME, "Choose an input PDF."); return
        self.run_btn.configure(state="disabled")
        def worker():
            try:
                asyncio.run(translate_pdf(
                    self.in_var.get(), self.out_var.get(),
                    target_lang=self.lang.get().strip() or "en",
                    dpi=int(self.dpi.get()),
                    close_browser=bool(self.close_chrome.get()),
                    log=self.log_write,
                    use_ocr=bool(self.use_ocr.get()),
                    ocr_lang=self.ocr_lang.get().strip() or "eng",
                ))
                self.log_write("✅ Yayy! We've Done It.Check for translated.pdf in your folder.")
            except Exception as e:
                self.log_write(f"❌ Error: {e}")
                messagebox.showerror(APP_NAME, str(e))
            finally:
                self.run_btn.configure(state="normal")
        Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    App().mainloop()
