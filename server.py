from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse
from pathlib import Path
from fastapi.responses import HTMLResponse
import shutil, uuid, asyncio
import uvicorn

from backup import translate_pdf

app = FastAPI(title="PDF Translator Bot 🌍")

UPLOADS = Path("uploads"); UPLOADS.mkdir(exist_ok=True)
OUTPUTS = Path("outputs"); OUTPUTS.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def home():
    return """
    <html>
      <head><title>PDF Translator Bot</title></head>
      <body style="font-family:Segoe UI; margin:40px;">
        <h1>🌍 PDF Translator Bot</h1>
        <form action="/translate/" enctype="multipart/form-data" method="post">
          <label>Upload PDF: <input type="file" name="file"></label><br><br>
          <label>Target Language: <input type="text" name="lang" value="en"></label><br><br>
          <button type="submit">Translate</button>
        </form>
      </body>
    </html>
    """

@app.post("/translate/")
async def translate(file: UploadFile, lang: str = Form("en")):
    file_id = uuid.uuid4().hex
    input_path = UPLOADS / f"{file_id}_{file.filename}"
    output_path = OUTPUTS / f"{file_id}_translated.pdf"

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    await translate_pdf(str(input_path), str(output_path), target_lang=lang, log=print)

    return FileResponse(output_path, filename=output_path.name)

if __name__ == "__main__":
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)