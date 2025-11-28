import os
import sys
import traceback

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# garante que o diretório desta função (api/) esteja no PYTHONPATH
sys.path.append(os.path.dirname(__file__))

from extractor import extract_table_from_xlsx  # noqa: E402

app = FastAPI(
    title="Extrator de Tabela do XLSX",
    description=(
        "Extrai ITEM, DESCRIÇÃO, UNID., QUANT., VALOR UNIT., VALOR TOTAL "
        "de uma planilha XLSX."
    ),
    version="1.0.0",
)

# CORS liberado para qualquer origem (ajuste se quiser restringir)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    return {"status": "ok", "message": "XLSX table extractor running"}


@app.post("/extract")
async def extract_table(file: UploadFile = File(...)):
    """
    Recebe um arquivo .xlsx e devolve um JSON no mesmo formato que o serviço antigo de PDF:

    {
      "rows": [
        {
          "item": int | null,
          "descricao": str,
          "unid": str,
          "quant": int | null,
          "valor_unit": float | null,
          "valor_total": float | null
        },
        ...
      ],
      "issues": [...]
    }
    """
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(
            status_code=400,
            detail="Somente arquivos XLSX (Excel moderno) são aceitos.",
        )

    try:
        file_bytes = await file.read()
        if not file_bytes:
            raise HTTPException(
                status_code=400,
                detail="Arquivo vazio.",
            )

        result = extract_table_from_xlsx(file_bytes)
        # o extrator já remove o cabeçalho; só retorna as linhas de itens
        return JSONResponse(content=result)

    except HTTPException:
        # repassa erros conhecidos
        raise
    except Exception as e:
        traceback_str = "".join(
            traceback.format_exception(type(e), e, e.__traceback__)
        )
        print("ERROR_PROCESSING_XLSX:", traceback_str)
        raise HTTPException(
            status_code=500,
            detail="Error processing XLSX",
        )
