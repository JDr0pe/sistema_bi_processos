import os
import sys
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ── Configuração ────────────────────────────────────────────────────────────
COLUNAS = {
    "id":          0,
    "processo":    2,
    "status":      9,
    "responsavel": 39,
    "data":        33,
}

# Descobre a pasta exata onde este script está rodando
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Cria o arquivo do banco de dados dentro desta mesma pasta
DB_PATH = os.path.join(BASE_DIR, "dados.db")


def selecionar_arquivo() -> str:
    """Abre o file dialog e retorna o caminho selecionado."""
    root = tk.Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    root.destroy()
    return arquivo


def ler_planilha(arquivo: str) -> pd.DataFrame:
    """Lê o Excel e valida número de colunas."""
    df = pd.read_excel(arquivo)
    max_idx = max(COLUNAS.values())
    if len(df.columns) <= max_idx:
        raise ValueError(f"Planilha tem apenas {len(df.columns)} colunas. Esperado mínimo {max_idx + 1}.")
    return df


# ── Transformação ────────────────────────────────────────────────────────────
def limpar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra por status finalizado ou prefixo TA023, deduplica, parseia datas e remove linhas com data inválida."""
    col_id       = df.columns[COLUNAS["id"]]
    col_processo = df.columns[COLUNAS["processo"]]
    col_status   = df.columns[COLUNAS["status"]]
    col_data     = df.columns[COLUNAS["data"]]

    # Filtrar apenas registros com status "finalizado" OU cujo processo inicia com "TA023"
    antes_filtro = len(df)
    cond_finalizado = df[col_status].astype(str).str.strip().str.lower() == "finalizado"
    cond_ta023      = df[col_processo].astype(str).str.strip().str.upper().str.startswith("TA023")
    df = df[cond_finalizado | cond_ta023]
    print(f"Filtro (finalizado ou TA023): {antes_filtro - len(df)} linha(s) filtrada(s), {len(df)} restantes.")

    antes = len(df)
    df = df.drop_duplicates(subset=[col_id])
    print(f"Deduplicação: {antes - len(df)} linha(s) removida(s), {len(df)} restantes.")

    df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
    nulas = df[col_data].isna().sum()
    if nulas:
        print(f"Atenção: {nulas} data(s) inválida(s) ignorada(s).")
    df = df.dropna(subset=[col_data])

    return df


def classificar(df: pd.DataFrame) -> pd.DataFrame:
    """Adiciona colunas Empresa, Ano e Mês."""
    col_processo = df.columns[COLUNAS["processo"]]
    col_data     = df.columns[COLUNAS["data"]]

    df["Empresa"] = df[col_processo].astype(str).str.upper().apply(
        lambda x: "CAP" if "CAP" in x else "SEG"
    )
    df["Ano"] = df[col_data].dt.year
    df["Mês"] = df[col_data].dt.month

    return df


# ── Persistência ─────────────────────────────────────────────────────────────
def get_conexao(db_path: str = DB_PATH) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    return sqlite3.connect(db_path)


def salvar_processos(df: pd.DataFrame, con: sqlite3.Connection) -> None:
    """Salva os processos novos na tabela 'processos'."""
    col_id = df.columns[COLUNAS["id"]]
    col_processo = df.columns[COLUNAS["processo"]]
    col_responsavel = df.columns[COLUNAS["responsavel"]]

    df_processos = pd.DataFrame({
        "id": df[col_id],
        "processo": df[col_processo],
        "responsavel_original": df[col_responsavel],
        "responsavel_atual": df[col_responsavel],
        "Empresa": df["Empresa"],
        "Ano": df["Ano"],
        "Mês": df["Mês"]
    })

    tabelas = pd.read_sql(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='processos'", con
    )

    if not tabelas.empty:
        ids_existentes = set(pd.read_sql("SELECT id FROM processos", con)["id"].astype(str).tolist())
        df_novos = df_processos[~df_processos["id"].astype(str).isin(ids_existentes)]
        if not df_novos.empty:
            df_novos.to_sql("processos", con, if_exists="append", index=False)
            print(f"Inseridos {len(df_novos)} novos registros na tabela 'processos'.")
        else:
            print("Nenhum novo registro para inserir.")
    else:
        df_processos.to_sql("processos", con, if_exists="append", index=False)
        print(f"Tabela 'processos' criada e {len(df_processos)} registros inseridos.")


# ── Pipeline ─────────────────────────────────────────────────────────────────
def processar_arquivo(arquivo: str, db_path: str = DB_PATH) -> None:
    """Pipeline completo: lê, transforma e persiste."""
    df   = ler_planilha(arquivo)
    df   = limpar_dados(df)
    df   = classificar(df)

    con = get_conexao(db_path)
    try:
        salvar_processos(df, con)
    finally:
        con.close()

    print(f"\nBanco de dados: {db_path}")


# ── Entrypoint ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    arquivo = selecionar_arquivo()
    if not arquivo:
        print("Nenhum arquivo selecionado.")
        sys.exit(0)

    try:
        processar_arquivo(arquivo)
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)