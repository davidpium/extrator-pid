from fastapi import FastAPI, UploadFile, File
import shutil
import os
from processador import processar_pdf
from fastapi.responses import FileResponse

app = FastAPI()

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.get("/")
def home():
    return {"status": "API rodando"}


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    caminho_pdf = os.path.join(UPLOAD_FOLDER, file.filename)

    with open(caminho_pdf, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    caminho_excel = processar_pdf(caminho_pdf)

    if not caminho_excel:
        return {"erro": "Nenhum instrumento encontrado"}

    return FileResponse(caminho_excel, filename="resultado.xlsx")