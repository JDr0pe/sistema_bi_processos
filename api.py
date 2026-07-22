from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import sqlite3, shutil, tempfile, os, pandas as pd

from funcoes import processar_arquivo, get_conexao

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # restrinja em produção
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve o front
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def root():
    return FileResponse("static/index.html")


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    """Recebe o Excel, processa e salva no banco."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Formato inválido. Envie .xlsx ou .xls")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    try:
        processar_arquivo(tmp_path)
    except Exception as e:
        raise HTTPException(500, str(e))
    finally:
        os.unlink(tmp_path)

    return {"ok": True, "mensagem": "Arquivo processado com sucesso."}


@app.get("/historico")
def historico():
    """Retorna toda a tabela histórico."""
    con = get_conexao()
    try:
        df = pd.read_sql("SELECT * FROM historico ORDER BY Ano, Mês", con)
        return df.to_dict(orient="records")
    except Exception:
        return []
    finally:
        con.close()


@app.get("/historico/resumo")
def resumo():
    """Totais por Tipo e Mês para o gráfico."""
    con = get_conexao()
    try:
        df = pd.read_sql("""
            SELECT Ano, Mês, Tipo, SUM(Quantidade) as Total
            FROM historico
            GROUP BY Ano, Mês, Tipo
            ORDER BY Ano, Mês
        """, con)
        return df.to_dict(orient="records")
    except Exception:
        return []
    finally:
        con.close()