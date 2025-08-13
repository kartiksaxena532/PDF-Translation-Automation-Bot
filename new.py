# bot.py
import os, time, subprocess, sys
from pathlib import Path
from mimetypes import guess_type
from urllib.request import urlopen
from urllib.error import URLError
import re
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ====== CONFIG ======
DEBUG_PORT = 9222
REMOTE = f"http://localhost:{DEBUG_PORT}"
# Common Chrome/Edge paths on Windows (adjust if needed)
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
USER_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", r"C:\Users\Public")) / "Chrome" / "PWProfile"
SRC_DIR = Path(r"H:\Bot\raw_images")
OUT_DIR = Path(r"H:\Bot\Translated Images"); OUT_DIR.mkdir(parents=True, exist_ok=True)
SOURCE_LANG = "auto"     # or "zh-CN"
TARGET_LANG = "en"       # change to "hi"/"fr"/etc for visible change
GLOB = ["*.png", "*.jpg", "*.jpeg", "*.webp", "*.bmp"]
# =====================

def debugger_ready() -> bool:
    try:
        with urlopen(f"{REMOTE}/json/version", timeout=1.5) as r:
            return r.status == 200
    except URLError:
        return False

def find_browser_exe() -> str:
    for p in CHROME_CANDIDATES + EDGE_CANDIDATES:
        if Path(p).is_file():
            return p
    raise FileNotFoundError("Chrome/Edge not found. Set CHROME_CANDIDATES/EDGE_CANDIDATES to your path.")

def launch_chrome_if_needed():
    if debugger_ready():
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
    # Start detached
    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
    # Wait until the debugger port is up
    for _ in range(60):  # ~15s
        if debugger_ready():
            return
        time.sleep(0.25)
    raise TimeoutError("Could not start Chrome with remote debugging port.")

def iter_images():
    files = []
    for pat in GLOB:
        files.extend(SRC_DIR.glob(pat))
    return sorted(files)

def upload_and_download_one(page, src_path: Path) -> Path:
    page.goto(f"https://translate.google.co.in/?sl={SOURCE_LANG}&tl={TARGET_LANG}&op=translate",
              wait_until="domcontentloaded")
    # Open Images tab
    page.get_by_role("button", name=re.compile(r"(Image translation|Images)", re.I)).click()

    # FilePayload (works even if H: is a mapped drive)
    payload = {
        "name": src_path.name,
        "mimeType": guess_type(src_path.name)[0] or "application/octet-stream",
        "buffer": src_path.read_bytes(),
    }

    # Upload
    try:
        page.locator('input[type="file"]').set_input_files(payload, timeout=4000)
    except Exception:
        with page.expect_file_chooser() as fc:
            page.get_by_role("button", name=re.compile(r"Browse your files", re.I)).click()
        fc.value.set_files(payload)

    # Wait for controls (the <img> may be hidden)
    page.get_by_role("button", name=re.compile(r"(Copy text|Show original|Show translated|Download)", re.I)).first.wait_for(timeout=60000)

    # Ensure translated overlay is ON: if "Show translated" is visible, you're on ORIGINAL → click once
    st = page.get_by_text("Show translated", exact=True)
    if st.count():
        st.first.click()

    # Prefer the official download
    out_path = OUT_DIR / f"{src_path.stem}_translated.png"
    try:
        with page.expect_download(timeout=15000) as dl_info:
            page.get_by_role("button", name=re.compile(r"(Download translation|Download)", re.I)).click()
        dl = dl_info.value
        ext = Path(dl.suggested_filename).suffix or ".png"
        out_path = OUT_DIR / f"{src_path.stem}_translated{ext}"
        dl.save_as(str(out_path))
        print(f"✓ Downloaded: {src_path.name} → {out_path.name}")
        return out_path
    except PWTimeout:
        pass
    except Exception as e:
        print(f"Download failed for {src_path.name}: {e} — using screenshot fallback.")

    # Screenshot fallback: panel that contains the viewer (captures overlay)
    panel = page.locator("div").filter(has=page.get_by_role("button", name=re.compile(r"Copy text", re.I))).first
    viewer = panel.locator(":scope div:has(canvas), :scope div:has(img.Jmlpdc)").last
    target = viewer if viewer.count() and viewer.is_visible() else panel
    try:
        target.scroll_into_view_if_needed()
    except Exception:
        pass
    page.wait_for_timeout(150)
    target.screenshot(path=str(out_path), animations="disabled")
    print(f"✓ Screenshotted: {src_path.name} → {out_path.name}")
    return out_path

def main():
    files = iter_images()
    if not files:
        print(f"No images found in {SRC_DIR}")
        return

    # 1) Ensure Chrome with remote debugging is running
    launch_chrome_if_needed()

    # 2) Attach and open a NEW TAB in that same window
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(REMOTE)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = ctx.new_page()  # new tab in the attached browser (new window if no context existed)

        ok = 0
        for i, src in enumerate(files, 1):
            try:
                upload_and_download_one(page, src)
                ok += 1
            except Exception as e:
                print(f"✗ Error on {src.name}: {e}")

        page.close()
        print(f"\nDone. {ok}/{len(files)} images saved to: {OUT_DIR}")

if __name__ == "__main__":
    main()
