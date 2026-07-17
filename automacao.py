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
    """Abre uma janela do explorador de arquivos para o usuário selecionar a planilha."""
    root = tk.Tk()
    root.withdraw() # Oculta a janela principal
    
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    root.destroy() # Encerra a instância do Tkinter adequadamente
    return arquivo

def processar_dados_excel(arquivo: str):
    """Lê a planilha e salva/atualiza os dados no banco de dados SQLite."""
    # ── Leitura ─────────────────────────────────────────────────────────────────
    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        sys.exit(1)

    # ── Validação de colunas ────────────────────────────────────────────────────
    max_idx = max(COLUNAS.values())
    if len(df.columns) <= max_idx:
        print(f"Planilha tem apenas {len(df.columns)} colunas. Esperado mínimo {max_idx + 1}.")
        sys.exit(1)

    col_id          = df.columns[COLUNAS["id"]]
    col_processo    = df.columns[COLUNAS["processo"]]
    col_responsavel = df.columns[COLUNAS["responsavel"]]
    col_data        = df.columns[COLUNAS["data"]]

    # ── Deduplicação ────────────────────────────────────────────────────────────
    antes = len(df)
    df = df.drop_duplicates(subset=[col_id]).copy()
    print(f"Deduplicação: {antes - len(df)} linha(s) removida(s), {len(df)} restantes.")

    # ── Parse de data ───────────────────────────────────────────────────────────
    df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")

    nulas = df[col_data].isna().sum()
    if nulas:
        print(f"Atenção: {nulas} data(s) inválida(s) ignorada(s).")
        df = df.dropna(subset=[col_data])

    # Se a planilha ficar vazia após remover datas nulas
    if df.empty:
        print("Erro: Não restaram dados válidos para processar.")
        sys.exit(1)

    # ── Classificação e campos derivados ────────────────────────────────────────
    df["Tipo"] = df[col_processo].astype(str).str.upper().apply(
        lambda x: "CAP" if "CAP" in x else "SEG"
    )
    df["Ano"] = df[col_data].dt.year
    df["Mês"] = df[col_data].dt.month

    # ── Conexão com o banco ──────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)

    try:
        # ── Tabela: consulta (snapshot da planilha) ────────────────────────
        # Renomeia colunas para nomes SQL-friendly antes de salvar
        df_consulta = df.rename(columns={
            col_id:          "id",
            col_processo:    "processo",
            col_responsavel: "responsavel",
            col_data:        "data",
        })

        ano_mes = df_consulta["data"].min().strftime("%Y%m")
        tabela_consulta = f"consulta_{ano_mes}"

        df_consulta.to_sql(tabela_consulta, con, if_exists="replace", index=False)
        print(f"Tabela '{tabela_consulta}' salva no banco.")

        # ── Tabela: histórico normalizado ────────────────────────────────────────────
        fato = (
            df.groupby([col_responsavel, "Tipo", "Ano", "Mês"])
            .size()
            .reset_index(name="Quantidade")
            .rename(columns={col_responsavel: "Responsavel"})
        )

        # Verifica se a tabela histórico já existe
        tabelas = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table' AND name='historico'", con)

        if not tabelas.empty:
            # Remove os períodos que serão reprocessados para evitar duplicação
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

        print(f"Tabela 'historico' atualizada no banco.")
    
    except Exception as e:
        print(f"Erro durante operações no banco de dados: {e}")
        con.rollback()
    finally:
        con.close()
        
    print(f"\nBanco de dados pronto em: {DB_PATH}")


def main():
    arquivo_selecionado = selecionar_arquivo()
    
    if not arquivo_selecionado:
        print("Operação cancelada. Nenhum arquivo selecionado.")
        sys.exit(0)
        
    print(f"Arquivo selecionado: {arquivo_selecionado}\nIniciando processamento...")
    processar_dados_excel(arquivo_selecionado)

if __name__ == "__main__":
    main()