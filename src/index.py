import os
import sys
import traceback

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

sys.path.append(os.path.dirname(__file__))

from extractor import extract_table_from_xlsx

app = FastAPI(
    title="Extrator de Tabela do XLSX",
    description="Extrai ITEM, DESCRIÇÃO, UNID., QUANT., VALOR UNIT., VALOR TOTAL de XLSX.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    return {"status": "ok"}


@app.post("/extract")
async def extract_table(file: UploadFile = File(...)):

    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(400, "Somente arquivos XLSX são aceitos.")

    try:
        content = await file.read()
        result = extract_table_from_xlsx(content)
        return JSONResponse(content=result)

    except Exception as e:
        print("ERROR_PROCESSING_XLSX:", e)
        raise HTTPException(500, "Error processing XLSX")
