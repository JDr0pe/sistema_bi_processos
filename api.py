from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
import sqlite3, shutil, tempfile, os, math, pandas as pd

from funcoes import processar_arquivo, get_conexao

class TransferenciaRequest(BaseModel):
    novo_responsavel: str

app = FastAPI()


def _sanitize_value(value):
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if pd.isna(value):
        return None
    return value


def _serializar_df(df: pd.DataFrame):
    """Converte DataFrame para estrutura JSON-safe."""
    if df.empty:
        return []

    records = []
    for _, row in df.iterrows():
        clean_row = {}
        for col in df.columns:
            clean_row[col] = _sanitize_value(row[col])
        records.append(clean_row)
    return records

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
    """Retorna toda a tabela processos."""
    con = get_conexao()
    try:
        df = pd.read_sql("""
            SELECT id, processo, responsavel_original, responsavel_atual AS Responsavel, Empresa, Ano, "Mês", 1 AS Quantidade
            FROM processos
            ORDER BY Ano DESC, "Mês" DESC, id DESC
        """, con)
        return _serializar_df(df)
    except Exception:
        return []
    finally:
        con.close()


@app.get("/historico/resumo")
def resumo():
    """Totais por Empresa, Mês e Responsável."""
    con = get_conexao()
    try:
        df = pd.read_sql("""
            SELECT Ano, "Mês", Empresa, responsavel_atual, COUNT(*) as Total
            FROM processos
            GROUP BY Ano, "Mês", Empresa, responsavel_atual
            ORDER BY Ano, "Mês"
        """, con)
        return df.to_dict(orient="records")
    except Exception:
        return []
    finally:
        con.close()


@app.patch("/processos/{processo_id}/transferir")
def transferir_processo(processo_id: str, request: TransferenciaRequest):
    con = get_conexao()
    try:
        cur = con.cursor()
        
        # Tenta atualizar assumindo que o ID pode ser texto
        cur.execute(
            "UPDATE processos SET responsavel_atual = ? WHERE id = ?",
            (request.novo_responsavel, processo_id)
        )
        con.commit()
        
        # Se nenhuma linha foi afetada e o ID é numérico, tenta converter
        if cur.rowcount == 0:
            try:
                numeric_id = float(processo_id)
                if numeric_id.is_integer():
                    numeric_id = int(numeric_id)
                cur.execute(
                    "UPDATE processos SET responsavel_atual = ? WHERE id = ?",
                    (request.novo_responsavel, numeric_id)
                )
                con.commit()
            except ValueError:
                pass
                
        if cur.rowcount == 0:
            raise HTTPException(status_code=404, detail="Processo não encontrado")
            
        return {"ok": True, "mensagem": f"Processo transferido para {request.novo_responsavel}"}
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        con.close()