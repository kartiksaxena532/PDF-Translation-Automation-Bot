import asyncio
import shutil
import time
import subprocess
from pathlib import Path
from typing import List, Optional
from mimetypes import guess_type
from urllib.request import urlopen
from urllib.error import URLError

import img2pdf
from PIL import Image
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ------------ Configuration (portable & safe) ------------
APP_NAME = "PDF Translator Bot - By Kartik Saxena"

def app_base_dir(drive_letter: str = None) -> Path:
    """
    Base storage directory (cross-platform).
    """
    if sys.platform.startswith("win"):
        drive_letter = drive_letter or "C:"
        base = Path(drive_letter) / "PDF_Translator_Bot"
        base.mkdir(parents=True, exist_ok=True)
        return base

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


# ------------ Chrome remote debugging attach ------------
DEBUG_PORT = 9222
REMOTE = f"http://localhost:{DEBUG_PORT}"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
EDGE_CANDIDATES = [
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]
USER_DATA_DIR = Path.home() / ".pw_chrome_profile"

IMG_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}


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
        log("Remote Chrome session detected — attaching…")
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
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
    log("Launching Chrome…")
    for _ in range(60):
        if dbg_ready():
            log("Attached to Chrome (remote debugging).")
            return
        time.sleep(0.25)
    raise TimeoutError("Could not start Chrome with remote debugging port.")


def wipe_images_only(folder: Path):
    if not folder.exists():
        return
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try:
                p.unlink(missing_ok=True)
            except:
                pass


# ------------ Core rendering: PyMuPDF ONLY ------------
def extract_pages(pdf_path: str, dpi: int, raw_dir: Path, log) -> List[Path]:
    """
    Render every page of the PDF to PNG using PyMuPDF.
    """
    raw_dir.mkdir(parents=True, exist_ok=True)
    for p in raw_dir.glob("*"):
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try:
                p.unlink(missing_ok=True)
            except:
                pass

    out_paths: List[Path] = []
    zoom = max(1.0, dpi / 72.0)
    mat = fitz.Matrix(zoom, zoom)

    with fitz.open(pdf_path) as doc:
        total = doc.page_count
        log(f"Rendering with PyMuPDF → {total} pages → {raw_dir}")
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=mat, alpha=False)  # RGB
            dst = raw_dir / f"page-{i:03}.png"
            pix.save(dst.as_posix())
            out_paths.append(dst)
            log(f"  ✓ {dst.name}")

    if not out_paths:
        raise RuntimeError("No images were produced by PyMuPDF rendering.")

    return out_paths


# ------------ DOCX writer ------------
class DocxLogger:
    def __init__(self, docx_path: Path, title: Optional[str] = None):
        self.path = Path(docx_path)
        self.doc = Document()

        # Set margins
        section = self.doc.sections[0]
        section.left_margin = Cm(2.12)
        section.right_margin = Cm(2.12)

        # Normal style
        normal = self.doc.styles["Normal"]
        nf = normal.font
        nf.name = "Segoe UI"
        nf.size = Pt(10)

        pf = normal.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.33

        if title:
            p = self.doc.add_heading(title, level=1)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            self._force_para_font(p)
            self.doc.add_paragraph("")

    def _force_para_font(self, para):
        for run in para.runs:
            run.font.name = "Segoe UI"
            run.font.size = Pt(10)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing = 1.33
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def add_section(self, heading: str, text: str):
        hp = self.doc.add_heading(heading, level=2)
        hp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self._force_para_font(hp)

        for line in text.splitlines():
            p = self.doc.add_paragraph(line if line.strip() else "")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            self._force_para_font(p)

        sp = self.doc.add_paragraph("")
        sp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self._force_para_font(sp)

    def save(self):
        for p in self.doc.paragraphs:
            if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            self._force_para_font(p)
        self.doc.save(self.path.as_posix())


# ------------ Translate via Google Translate Images (Playwright) ------------
async def translate_images(
    img_paths: List[Path],
    target_lang: str,
    close_browser: bool,
    trans_dir: Path,
    log=None,
    docx_logger: Optional[DocxLogger] = None,
) -> List[Path]:
    trans_dir.mkdir(parents=True, exist_ok=True)
    for p in trans_dir.glob("*"):
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            try:
                p.unlink(missing_ok=True)
            except:
                pass

    results: List[Path] = []
    url = f"https://translate.google.co.in/?sl=auto&tl={target_lang}&op=images"
    launch_chrome_if_needed(log or (lambda *_: None))

    async with async_playwright() as pw:
        browser = await pw.chromium.connect_over_cdp(REMOTE)
        try:
            ctx = browser.contexts[0] if browser.contexts else await browser.new_context(accept_downloads=True)
            page = await ctx.new_page()
            await page.goto(url)

            download_btn = page.get_by_role("button", name="Download translation")
            copy_btn_css = page.locator('button[aria-label="Copy text"]')

            for img in img_paths:
                log(f"Processing {img.name}...")
                await page.locator('input[type="file"]').set_input_files(str(img))

                await download_btn.first.wait_for(state="visible", timeout=60000)
                async with page.expect_download() as dl_info:
                    await download_btn.first.click()
                dl = await dl_info.value
                out_img = trans_dir / f"{img.stem}-translated.png"
                await dl.save_as(str(out_img))
                results.append(out_img)

                copied = ""
                try:
                    if await copy_btn_css.count():
                        await copy_btn_css.first.click()
                        await page.wait_for_timeout(150)
                        copied = await page.evaluate("navigator.clipboard.readText()")
                except Exception:
                    copied = ""

                if copied and docx_logger:
                    docx_logger.add_section(img.name, copied.strip())
                    log(f"Added text for {img.name} to DOCX")

            await page.close()
        finally:
            if close_browser:
                try:
                    cdp = await browser.new_browser_cdp_session()
                    await cdp.send("Browser.close")
                except:
                    pass
                await browser.close()

    return results


# ------------ PDF Builder ------------
def build_pdf(images: List[Path], out_pdf: Path, log):
    log("Building final PDF…")
    normalized: List[Path] = []
    for p in images:
        if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
            normalized.append(p)
        else:
            with Image.open(p) as im:
                q = p.with_suffix(".png")
                im.save(q)
                normalized.append(q)
    with open(out_pdf, "wb") as f:
        f.write(img2pdf.convert([str(p) for p in normalized]))
    log(f"✓ Saved → {out_pdf}")


# ------------ Main Translator ------------
async def translate_pdf(input_pdf: str, output_pdf: str, target_lang: str = "en", dpi: int = 150, close_browser: bool = True, log=print):
    if not Path(input_pdf).is_file():
        raise FileNotFoundError(f"File not found: {input_pdf}")

    run_dir = new_run_dir()
    raw_dir = run_dir / "raw"
    trans_dir = run_dir / "translated"

    try:
        Path(output_pdf).parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        fallback = app_base_dir() / Path(output_pdf).name
        log(f"⚠ Output folder not available. Using: {fallback}")
        output_pdf = str(fallback)

    docx_path = Path(output_pdf).with_suffix(".docx")
    docx_logger = DocxLogger(docx_path, title="Translated Text")

    pages = extract_pages(input_pdf, dpi, raw_dir, log)
    translated = await translate_images(pages, target_lang, close_browser, trans_dir, log, docx_logger)
    build_pdf(translated, Path(output_pdf), log)

    try:
        docx_logger.save()
        log(f"✓ Saved DOCX → {docx_path}")
    except Exception as e:
        log(f"⚠ Could not save DOCX: {e}")

    wipe_images_only(raw_dir)
    wipe_images_only(trans_dir)
    log(f"Cleaned up temporary images in {run_dir}.")
