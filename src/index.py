import os
import sys
import traceback

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# garante que o diretório desta função (api/) esteja no PYTHONPATH
sys.path.append(os.path.dirname(__file__))

from extractor import extract_table_from_pdf  # noqa: E402

app = FastAPI(
    title="Extrator de Tabela do PDF",
    description="Extrai ITEM, DESCRIÇÃO, UNID., QUANT., VALOR UNIT., VALOR TOTAL de um PDF.",
    version="1.2.1",
)

# CORS liberado para qualquer origem (em produção, restringir)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "Use POST /api/index/extract com campo 'file' para enviar o PDF.",
        "endpoints": {
            "health": "/api/index/health",
            "extract": "/api/index/extract",
        },
    }


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    # valida extensão
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="Empty file")

        # chama o parser principal
        result = extract_table_from_pdf(content)  # retorna {"rows": [...], "issues": [...]}
        rows = result.get("rows", [])
        issues = result.get("issues", [])

        # métricas agregadas úteis pro SIGECON
        total_itens = sum(1 for r in rows if r.get("item") is not None)
        soma_total = round(sum((r.get("valor_total") or 0.0) for r in rows), 2)
        soma_unitaria = round(sum((r.get("valor_unit") or 0.0) for r in rows), 2)

        return JSONResponse(
            {
                "count": len(rows),
                "count_com_item": total_itens,
                "soma_valor_total": soma_total,
                "soma_valor_unit": soma_unitaria,
                "moeda": "BRL",
                "columns": [
                    "item",
                    "descricao",
                    "unid",
                    "quant",
                    "valor_unit",
                    "valor_total",
                ],
                "rows": rows,
                "issues": issues,
            }
        )

    except HTTPException:
        # repassa erros conhecidos
        raise
    except Exception as e:
        traceback_str = "".join(
            traceback.format_exception(type(e), e, e.__traceback__)
        )
        print("ERROR_PROCESSING_PDF:", traceback_str)
        raise HTTPException(
            status_code=500,
            detail="Error processing PDF",
        )
