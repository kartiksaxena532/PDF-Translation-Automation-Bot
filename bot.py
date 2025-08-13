# translate_pdf_via_gui.py  – Aug 2025
# - Starts (or attaches to) Chrome via CDP so a new TAB opens in your window
# - PDF -> page PNGs in RAW_DIR
# - Uploads each page to Google Translate (Images), waits 3s, downloads translated image
# - Builds a single PDF from the translated images
# - Wipes PNG/JPG/etc. from RAW_DIR and TRANS_DIR and CLOSES the browser window

import asyncio, os, pathlib, img2pdf, sys, shutil, time, subprocess, re
from pathlib import Path
from mimetypes import guess_type
from urllib.request import urlopen
from urllib.error import URLError

from pdf2image import convert_from_path, pdfinfo_from_path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout
from dotenv import load_dotenv
load_dotenv()

# ====== FOLDERS / PATHS ======
RAW_DIR   = pathlib.Path(r"H:\Bot\raw images")
TRANS_DIR = pathlib.Path(r"H:\Bot\translated images")
POPPLER   = pathlib.Path(
    r"H:\Downloads\Oppo Amos EP040\Release-24.08.0-0\poppler-24.08.0\Library\bin"
)
URL = "https://translate.google.co.in/?sl=auto&tl=en&op=images"

RAW_DIR.mkdir(parents=True, exist_ok=True)
TRANS_DIR.mkdir(parents=True, exist_ok=True)

# ====== CHROME REMOTE DEBUG / ATTACH ======
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

def _debugger_ready() -> bool:
    try:
        with urlopen(f"{REMOTE}/json/version", timeout=1.5) as r:
            return r.status == 200
    except URLError:
        return False

def _find_browser_exe() -> str:
    for p in CHROME_CANDIDATES + EDGE_CANDIDATES:
        if Path(p).is_file():
            return p
    raise FileNotFoundError("Chrome/Edge not found. Update CHROME_CANDIDATES/EDGE_CANDIDATES.")

def _launch_chrome_if_needed():
    if _debugger_ready():
        return
    exe = _find_browser_exe()
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
    for _ in range(60):
        if _debugger_ready():
            return
        time.sleep(0.25)
    raise TimeoutError("Could not start Chrome with remote debugging port.")

# ====== HELPERS ======
IMG_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}

def wipe_images_only(folder: Path):
    """Delete only image files in folder (keep PDFs and anything else)."""
    if not folder.exists():
        return
    for p in folder.iterdir():
        try:
            if p.is_file() and p.suffix.lower() in IMG_EXTS:
                p.unlink(missing_ok=True)
        except Exception:
            pass

# ====== PDF → IMAGES ======
def extract_pages(pdf_path: str, dpi: int = 150) -> list[str]:
    info = pdfinfo_from_path(pdf_path, poppler_path=str(POPPLER))
    print(f"🔍  {info['Pages']} pages – saving PNGs into {RAW_DIR}")
    # fresh RAW_DIR
    shutil.rmtree(RAW_DIR, ignore_errors=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    paths = []
    for i, pg in enumerate(
        convert_from_path(pdf_path, dpi=dpi, fmt="png", poppler_path=str(POPPLER)), 1
    ):
        out = RAW_DIR / f"page-{i:03}.png"
        pg.save(out)
        paths.append(str(out))
        print(f"   ✓ {out.name}")
    return paths

# ====== TRANSLATE EACH IMAGE VIA GOOGLE TRANSLATE (Images) ======
async def translate_images(img_paths: list[str]) -> list[str]:
    # fresh TRANS_DIR
    shutil.rmtree(TRANS_DIR, ignore_errors=True)
    TRANS_DIR.mkdir(parents=True, exist_ok=True)
    translated: list[str] = []

    # Ensure Chrome with remote debugging is running
    _launch_chrome_if_needed()

    async with async_playwright() as p:
        # Attach to that Chrome and open a NEW TAB in the same window
        browser = await p.chromium.connect_over_cdp(REMOTE)
        try:
            ctx = browser.contexts[0] if browser.contexts else await browser.new_context(accept_downloads=True)
            page = await ctx.new_page()
            await page.goto(URL)

            browse_btn   = page.get_by_role("button", name="Browse your files")
            download_btn = page.get_by_role("button", name=re.compile(r"Download translation|Download", re.I))
            clear_btn    = page.get_by_role("button", name=re.compile(r"Clear image|Clear", re.I))

            for idx, img in enumerate(img_paths, 1):
                print(f"🌐  Translating {idx}/{len(img_paths)} …")

                # Choose file (prefer direct input with FilePayload)
                payload = {
                    "name": Path(img).name,
                    "mimeType": guess_type(Path(img).name)[0] or "application/octet-stream",
                    "buffer": Path(img).read_bytes(),
                }
                try:
                    await page.locator('input[type="file"]').set_input_files(payload, timeout=2000)
                except Exception:
                    async with page.expect_file_chooser() as fc_info:
                        await browse_btn.click()
                    chooser = await fc_info.value
                    await chooser.set_files(payload)

                # Wait for Download button to appear
                await download_btn.first.wait_for(state="visible", timeout=60000)

                # Click Download & capture file
                async with page.expect_download() as dl_info:
                    await download_btn.first.click()
                dl_file = await dl_info.value
                suggested = dl_file.suggested_filename or f"{Path(img).stem}-translated.png"
                ext = Path(suggested).suffix or ".png"
                out_png = TRANS_DIR / f"{Path(img).stem}-translated{ext}"
                await dl_file.save_as(out_png)
                translated.append(str(out_png))
                print(f"   ✓ saved {out_png.name}")

                # Clear canvas if the button exists
                try:
                    await clear_btn.click()
                except Exception:
                    pass

            await page.close()
        finally:
            # 🔚 Close the attached Chrome window itself
            try:
                await browser.close()
            except Exception:
                pass

    return translated

# ====== IMAGES → SINGLE PDF ======
def build_pdf(img_paths: list[str], pdf_out: str):
    print("📚  Building final PDF …")
    with open(pdf_out, "wb") as f:
        f.write(img2pdf.convert([open(p, "rb").read() for p in img_paths]))
    print(f"🎉  Saved → {pdf_out}")

# ====== MAIN ======
async def main(src_pdf: str, out_pdf: str):
    if not os.path.isfile(src_pdf):
        sys.exit(f"❌ File not found: {src_pdf}")

    raw_imgs   = extract_pages(src_pdf, dpi=150)
    trans_imgs = await translate_images(raw_imgs)
    build_pdf(trans_imgs, out_pdf)

    # 🧹 wipe ONLY images in both folders (keep your resulting PDF even if it’s in TRANS_DIR)
    wipe_images_only(RAW_DIR)
    wipe_images_only(TRANS_DIR)
    print("🧹 Cleaned up images in RAW and TRANSLATED folders.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(
        description="Translate a PDF via Google-Translate Images tab (CDP attach), then clean up images."
    )
    parser.add_argument("input_pdf",  help="Path to source PDF")
    parser.add_argument("output_pdf", nargs="?", default="translated.pdf",
                        help="Filename for translated PDF")
    args = parser.parse_args()

    asyncio.run(main(args.input_pdf, args.output_pdf))
