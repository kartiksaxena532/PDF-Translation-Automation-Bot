import sys, os, re, time, subprocess, asyncio, shutil
from urllib.request import urlopen
from urllib.error import URLError
from pathlib import Path
from mimetypes import guess_type
from threading import Thread
from typing import List, Optional

import img2pdf
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image, ImageStat
try:
    import fitz  # PyMuPDF fallback renderer
except Exception:
    fitz = None

from playwright.async_api import async_playwright, TimeoutError as PWTimeout
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ------------ Configuration (portable & safe) ------------
APP_NAME = "PDF Translator Bot - By Kartik Saxena"

def app_base_dir(drive_letter: str = None) -> Path:
    """
    Allow the user to manually specify their preferred drive for storage.
    User can fill in the drive letter part themselves. 
    """
    if sys.platform.startswith("win"):
        # Use user input or default to C:
        drive_letter = drive_letter or "C:"
        base = Path(drive_letter) / "PDF_Translator_Bot"
        base.mkdir(parents=True, exist_ok=True)
        return base

    # For non-Windows systems
    root = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    base = root / "PDF_Translator_Bot"
    base.mkdir(parents=True, exist_ok=True)
    return base

def new_run_dir() -> Path:
    """Unique per run; avoids wiping anything outside this isolated sandbox."""
    from datetime import datetime
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    d = app_base_dir() / f"run-{stamp}"
    (d / "raw").mkdir(parents=True, exist_ok=True)
    (d / "translated").mkdir(parents=True, exist_ok=True)
    return d

# POPPLER location resolver: use bundled folder in EXE, else env var / common paths
def resolve_poppler_bin() -> Path:
    # In bundled EXE (PyInstaller)
    if getattr(sys, "_MEIPASS", None):
        p = Path(sys._MEIPASS) / "poppler_bin"
        if p.exists():
            return p

    # Dev/installed fallbacks (custom dev path first)
    candidates = [
        Path(r"H:\Bot\poppler-24.08.0\Library\bin"),   # DEV PATH (custom)
        Path(os.environ.get("POPPLLER_BIN", "")),      # common typo guard
        Path(os.environ.get("POPPLER_BIN", "")),
        Path(r"C:\Program Files\poppler\bin"),
        Path(r"C:\Program Files (x86)\poppler\bin"),
        Path("/usr/local/bin"),
        Path("/usr/bin"),
    ]
    for c in candidates:
        if c and c.exists():
            return c

    raise FileNotFoundError(
        "Poppler not found. Bundle 'poppler_bin' with the EXE, "
        "set POPPLER_BIN env var, or install Poppler."
    )

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
    except Exception:
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
        log("Chrome would be starting any minute soon please wait....")
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

# ---- Render validation & MuPDF fallback (fixes "blank PNGs" from OCR'd PDFs) ----
def is_mostly_blank(img_path: Path, thresh: float = 0.995) -> bool:
    try:
        with Image.open(img_path) as im:
            im = im.convert("L")
            stat = ImageStat.Stat(im)
            mean = stat.mean[0]
            var = stat.var[0]
            if mean > 250 and var < 5:
                small = im.resize((max(1, im.width // 16), max(1, im.height // 16)))
                pixels = list(small.getdata())
                nonwhite = sum(1 for p in pixels if p < 248)
                ratio = nonwhite / len(pixels)
                return ratio < (1 - thresh)
            return False
    except Exception:
        return False

def render_with_pymupdf(pdf_path: str, dpi: int, out_dir: Path, log) -> List[Path]:
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) not installed. pip install pymupdf")
    out_paths: List[Path] = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    log("Re-rendering with MuPDF (fallback)…")
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=mat, alpha=False)
            dst = out_dir / f"page-{i:03}.png"
            pix.save(dst.as_posix())
            out_paths.append(dst)
            log(f"  ✓ (MuPDF) {dst.name}")
    return out_paths

# ------------ Core pipeline ------------
def extract_pages(pdf_path: str, dpi: int, raw_dir: Path, log) -> List[Path]:
    """First try Poppler (pdf2image). If pages look blank, fallback to MuPDF (PyMuPDF)."""
    poppler_bin = resolve_poppler_bin()
    info = pdfinfo_from_path(pdf_path, poppler_path=str(poppler_bin))
    log(f"Processing your PDF which has {info['Pages']} pages → saving PNGs in {raw_dir}")

    raw_dir.mkdir(parents=True, exist_ok=True)
    for p in raw_dir.glob("*"):
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try: p.unlink(missing_ok=True)
            except: pass

    out_paths: List[Path] = []
    try:
        pages = convert_from_path(pdf_path, dpi=dpi, fmt="png", poppler_path=str(poppler_bin))
        for i, pg in enumerate(pages, 1):
            out = raw_dir / f"page-{i:03}.png"
            pg.save(out)
            out_paths.append(out)
            log(f"  ✓ {out.name}")
    except Exception as e:
        log(f"⚠ Poppler render failed: {e}")

    need_fallback = False
    if out_paths:
        sample = out_paths[:min(3, len(out_paths))]
        if len(out_paths) > 3:
            sample.append(out_paths[-1])
        blanks = sum(1 for p in sample if is_mostly_blank(p))
        if blanks >= max(2, len(sample) // 2):
            need_fallback = True
            log("⚠ Detected mostly blank pages from Poppler render; will fallback to MuPDF.")
    else:
        need_fallback = True
        log("⚠ No pages produced by Poppler; will fallback to MuPDF.")

    if need_fallback:
        for p in out_paths:
            try: p.unlink(missing_ok=True)
            except: pass
        out_paths = render_with_pymupdf(pdf_path, dpi, raw_dir, log)

    return out_paths

async def translate_images(
    img_paths: List[Path],
    target_lang: str,
    close_browser: bool,
    trans_dir: Path,
    log=None,
    txt_append_path: Optional[Path] = None,   # NEW
) -> List[Path]:
    """
    Upload images to Google Translate, download translated images,
    and (NEW) copy translated text to translated.txt.
    """
    trans_dir.mkdir(parents=True, exist_ok=True)
    for p in trans_dir.glob("*"):
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try: p.unlink(missing_ok=True)
            except: pass

    results: List[Path] = []
    url = f"https://translate.google.co.in/?sl=auto&tl={target_lang}&op=images"
    launch_chrome_if_needed(log or (lambda *_: None))

    async with async_playwright() as pw:
        browser = await pw.chromium.connect_over_cdp(REMOTE)
        try:
            ctx = browser.contexts[0] if browser.contexts else await browser.new_context(accept_downloads=True)

            # Grant clipboard permissions so Copy text works reliably
            origin = "https://translate.google.co.in"
            try:
                await ctx.grant_permissions(["clipboard-read", "clipboard-write"], origin=origin)
            except Exception:
                pass

            page = await ctx.new_page()
            await page.goto(url)

            browse_btn   = page.get_by_role("button", name=re.compile(r"Browse your files", re.I))
            download_btn = page.get_by_role("button", name=re.compile(r"(Download translation|Download)", re.I))
            clear_btn    = page.get_by_role("button", name=re.compile(r"(Clear image|Clear)", re.I))
            show_translated = page.get_by_text("Show translated", exact=True)
            copy_btn_role  = page.get_by_role("button", name=re.compile(r"Copy text", re.I))
            copy_btn_css   = page.locator('button[aria-label="Copy text"]')  # fallback

            total = len(img_paths)
            for idx, img in enumerate(img_paths, 1):
                log(f"[{idx}/{total}] {img.name}")

                payload = {
                    "name": img.name,
                    "mimeType": guess_type(img.name)[0] or "application/octet-stream",
                    "buffer": img.read_bytes()
                }
                try:
                    await page.locator('input[type="file"]').set_input_files(payload, timeout=1500)
                except Exception:
                    async with page.expect_file_chooser() as fc:
                        await browse_btn.click()
                    chooser = await fc.value
                    await chooser.set_files(payload)

                if await show_translated.count() > 0:
                    await show_translated.first.click()

                await download_btn.first.wait_for(state="visible", timeout=60000)
                async with page.expect_download() as dl_info:
                    await download_btn.first.click()
                dl = await dl_info.value
                suggested = dl.suggested_filename or f"{img.stem}-translated.png"
                ext = Path(suggested).suffix or ".png"
                out_img = trans_dir / f"{img.stem}-translated{ext}"
                await dl.save_as(str(out_img))
                results.append(out_img)
                log(f"   ↳ saved {out_img.name}")

                # --- NEW: Copy text and append to file
                copied = ""
                try:
                    target_btn = copy_btn_role if await copy_btn_role.count() else copy_btn_css
                    if await target_btn.count():
                        await target_btn.first.click()
                        await page.wait_for_timeout(150)
                        copied = await page.evaluate("navigator.clipboard.readText()")
                except Exception:
                    copied = ""

                if copied and txt_append_path:
                    try:
                        with open(txt_append_path, "a", encoding="utf-8") as fp:
                            fp.write(f"\n===== {img.name} =====\n")
                            fp.write(copied.strip())
                            fp.write("\n")
                        log(f"   ↳ appended text for {img.name} to {txt_append_path.name}")
                    except Exception as e:
                        log(f"⚠ Could not append text for {img.name}: {e}")

                # Clear for next image
                try: await clear_btn.click()
                except: pass

            await page.close()
        finally:
            if close_browser:
                # Try a graceful browser-wide close via DevTools
                try:
                    # Ask the debugging browser to exit entirely
                    cdp = await browser.new_browser_cdp_session()
                    await cdp.send("Browser.close")
                except Exception:
                    pass

                # Close Playwright connection (safe even if Browser.close already ended it)
                try:
                    await browser.close()
                except Exception:
                    pass

                # If we launched Chrome ourselves, kill the whole process tree as fallback
              
    return results

def build_pdf(images: List[Path], out_pdf: Path, log):
    log("Building Your Final PDF…")
    normalized: List[Path] = []
    for p in images:
        if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
            normalized.append(p); continue
        with Image.open(p) as im:
            q = p.with_suffix(".png"); im.save(q); normalized.append(q)
    with open(out_pdf, "wb") as f:
        f.write(img2pdf.convert([str(p) for p in normalized]))
    log(f"✓ Saved → {out_pdf}")

async def translate_pdf(input_pdf: str, output_pdf: str, target_lang: str = "en", dpi: int = 150, close_browser: bool = True, log=print):
    if not Path(input_pdf).is_file():
        raise FileNotFoundError(f"File not found: {input_pdf}")

    run_dir = new_run_dir()
    raw_dir = run_dir / "raw"
    trans_dir = run_dir / "translated"

    # Ensure output directory exists / recover if invalid
    try:
        Path(output_pdf).parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        fallback = app_base_dir() / Path(output_pdf).name
        log(f"⚠ Output folder not available. Using: {fallback}")
        output_pdf = str(fallback)

    # NEW: create text file path next to output PDF
    txt_append_path = Path(output_pdf).with_name("translated.txt")
    # If you prefer a fresh file per run, uncomment next lines:
    # try: txt_append_path.unlink(missing_ok=True)
    # except: pass

    pages = extract_pages(input_pdf, dpi, raw_dir, log)
    translated = await translate_images(
        pages, target_lang, close_browser, trans_dir, log,
        txt_append_path=txt_append_path,   # NEW
    )
    build_pdf(translated, Path(output_pdf), log)

    wipe_images_only(raw_dir)
    wipe_images_only(trans_dir)
    log(f"Cleaned up temporary images in {run_dir}.")

# ------------ GUI ------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x380")
        self.resizable(False, False)

        self.in_var  = tk.StringVar()
        self.out_var = tk.StringVar(value="")  # Empty output path by default
        self.lang    = tk.StringVar(value="en")
        self.dpi     = tk.IntVar(value=150)
        self.close_chrome = tk.BooleanVar(value=True)

        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)

        # Input PDF section
        ttk.Label(frm, text="Input PDF:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.in_var, width=64).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Browse…", command=self.pick_in).grid(row=0, column=2)

        # Output PDF section (empty path by default)
        ttk.Label(frm, text="Output PDF:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.out_var, width=64).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Choose…", command=self.pick_out).grid(row=1, column=2)

        # Target language and DPI section
        opt = ttk.Frame(frm); opt.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8,4))
        ttk.Label(opt, text="Target language (tl):").grid(row=0, column=0, sticky="w")
        ttk.Entry(opt, textvariable=self.lang, width=8).grid(row=0, column=1, padx=6)
        ttk.Label(opt, text="DPI:").grid(row=0, column=2, sticky="w")
        ttk.Entry(opt, textvariable=self.dpi, width=6).grid(row=0, column=3, padx=6)
        ttk.Checkbutton(opt, text="Close Chrome on finish", variable=self.close_chrome).grid(row=0, column=4, padx=12)

        # Buttons to start and quit
        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=3, pady=6, sticky="ew")
        self.run_btn = ttk.Button(btns, text="Translate PDF", command=self.start)
        self.run_btn.pack(side="left")
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="right")

        # Log display section
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
        if not self.out_var.get():
            messagebox.showwarning(APP_NAME, "Specify an output path."); return
        self.run_btn.configure(state="disabled")
        def worker():
            try:
                asyncio.run(translate_pdf(
                    self.in_var.get(), self.out_var.get(),
                    target_lang=self.lang.get().strip() or "en",
                    dpi=int(self.dpi.get()),
                    close_browser=bool(self.close_chrome.get()),
                    log=self.log_write
                ))
                self.log_write("✅ Yayy! We've Done It. Check your output PDF.")
            except Exception as e:
                self.log_write(f"❌ Error: {e}")
                messagebox.showerror(APP_NAME, str(e))
            finally:
                self.run_btn.configure(state="normal")
        Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    App().mainloop()