To be installed - 
import asyncio, os, pathlib, img2pdf, sys, shutil, time, subprocess, re
from pathlib import Path
from mimetypes import guess_type
from urllib.request import urlopen
from urllib.error import URLError

from pdf2image import convert_from_path, pdfinfo_from_path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout
from dotenv import load_dotenv

Path for  -  POPPLER   = pathlib.Path(
    r"H:\Downloads\Oppo Amos EP040\Release-24.08.0-0\poppler-24.08.0\Library\bin"
)

python bot.py "H:\Bot\undoc.pdf" "H:\Bot\translated.pdf"  
Clears all the folders after PDF creation


$POPPLER_BIN = "H:\Downloads\Oppo Amos EP040\Release-24.08.0-0\poppler-24.08.0\Library\bin"   # <-- change to your path

pyinstaller Kartik_translation_bot.py `
  --onefile `
  --noconsole `
  --collect-all playwright `
  --add-data "$POPPLER_BIN;poppler_bin"


# in your project folder
>> py -3.11 -m venv .venv
>> .\.venv\Scripts\Activate.ps1
>> python -m pip install --upgrade pip
>> pip install pyinstaller playwright pdf2image img2pdf pillow

>>



Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
>> .\.venv\Scripts\Activate.ps1


setx PYLAUNCHER_ALLOW_INSTALL 1
>> py -3.11 -V    # triggers install of Python 3.11 if missing


 winget install --id Python.Python.3.11 -e
>> # (or a newer one)
>> # winget install --id Python.Python.3.12 -e

where py
>> py -0     # that's a zero → lists all detected Pythons
>> where python
>> python --version


# Chrome     
>> & "C:\Program Files\Google\Chrome\Application\chrome.exe" `
>>   --remote-debugging-port=9222 `
>>   --user-data-dir="$env:LOCALAPPDATA\Chrome\PWProfile"


