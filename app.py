from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os
import uuid
from processador import processar_pdf

app = FastAPI()

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    
    nome_unico = str(uuid.uuid4()) + ".pdf"
    caminho_pdf = os.path.join(UPLOAD_FOLDER, nome_unico)

    with open(caminho_pdf, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    caminho_excel = processar_pdf(caminho_pdf)

    return FileResponse(
        path=caminho_excel,
        filename="resultado.xlsx",
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
