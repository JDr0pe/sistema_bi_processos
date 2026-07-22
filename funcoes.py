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
    "responsavel": 28,
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
    """Deduplica, parseia datas e remove linhas com data inválida."""
    col_id   = df.columns[COLUNAS["id"]]
    col_data = df.columns[COLUNAS["data"]]

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
    """Adiciona colunas Tipo, Ano e Mês."""
    col_processo = df.columns[COLUNAS["processo"]]
    col_data     = df.columns[COLUNAS["data"]]

    df["Tipo"] = df[col_processo].astype(str).str.upper().apply(
        lambda x: "CAP" if "CAP" in x else "SEG"
    )
    df["Ano"] = df[col_data].dt.year
    df["Mês"] = df[col_data].dt.month

    return df


def gerar_fato(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega dados para a tabela fato (histórico normalizado)."""
    col_responsavel = df.columns[COLUNAS["responsavel"]]

    return (
        df.groupby([col_responsavel, "Tipo", "Ano", "Mês"])
        .size()
        .reset_index(name="Quantidade")
        .rename(columns={col_responsavel: "Responsavel"})
    )


# ── Persistência ─────────────────────────────────────────────────────────────
def get_conexao(db_path: str = DB_PATH) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    return sqlite3.connect(db_path)


def salvar_consulta(df: pd.DataFrame, con: sqlite3.Connection) -> str:
    """Salva snapshot completo da planilha. Retorna o nome da tabela criada."""
    col_data = df.columns[COLUNAS["data"]]
    ano_mes  = df[col_data].min().strftime("%Y%m")

    df_consulta = df.rename(columns={
        df.columns[COLUNAS["id"]]:          "id",
        df.columns[COLUNAS["processo"]]:    "processo",
        df.columns[COLUNAS["responsavel"]]: "responsavel",
        col_data:                           "data",
    })

    tabela = f"consulta_{ano_mes}"
    df_consulta.to_sql(tabela, con, if_exists="replace", index=False)
    print(f"Tabela '{tabela}' salva no banco.")
    return tabela


def salvar_historico(fato: pd.DataFrame, con: sqlite3.Connection) -> None:
    """Upsert no histórico: remove períodos afetados e insere novos dados."""
    tabelas = pd.read_sql(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='historico'", con
    )

    if not tabelas.empty:
        periodos = fato[["Responsavel", "Tipo", "Ano", "Mês"]].drop_duplicates()
        for _, row in periodos.iterrows():
            con.execute(
                "DELETE FROM historico WHERE Responsavel=? AND Tipo=? AND Ano=? AND Mês=?",
                (row["Responsavel"], row["Tipo"], row["Ano"], row["Mês"])
            )
        con.commit()
        fato.to_sql("historico", con, if_exists="append", index=False)
    else:
        fato.to_sql("historico", con, if_exists="replace", index=False)

    print("Tabela 'historico' atualizada no banco.")


# ── Pipeline ─────────────────────────────────────────────────────────────────
def processar_arquivo(arquivo: str, db_path: str = DB_PATH) -> None:
    """Pipeline completo: lê, transforma e persiste."""
    df   = ler_planilha(arquivo)
    df   = limpar_dados(df)
    df   = classificar(df)
    fato = gerar_fato(df)

    con = get_conexao(db_path)
    try:
        salvar_consulta(df, con)
        salvar_historico(fato, con)
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