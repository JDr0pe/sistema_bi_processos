import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo() -> str:
    """Abre uma janela do explorador de arquivos para o usuário selecionar a planilha."""
    root = tk.Tk()
    root.withdraw()
    
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    return arquivo

def processar_dados_excel(arquivo: str):
    """Lê, processa os dados da planilha e salva o histórico atualizado."""
    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        sys.exit(1)

    # ── Mapeamento de Colunas (Modifique aqui se os nomes mudarem) ──────────
    # Usar nomes é mais seguro, mas se precisar manter por índice:
    try:
        col_id          = df.columns[0]
        col_processo    = df.columns[2]
        col_responsavel = df.columns[28]
        col_data        = df.columns[33]
    except IndexError:
        print("Erro: A planilha não possui a quantidade de colunas esperada.")
        sys.exit(1)

    # ── Deduplicação ────────────────────────────────────────────────────────────
    df = df.drop_duplicates(subset=[col_id]).copy()

    # ── Data mais antiga → nome do arquivo de consulta ─────────────────────────
    # Coerção de erros ajuda a evitar travamentos com dados sujos
    df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
    data_mais_antiga = df[col_data].min()
    
    if pd.isna(data_mais_antiga):
        print("Erro: Nenhuma data válida encontrada na coluna especificada.")
        sys.exit(1)
        
    ano_mes = data_mais_antiga.strftime("%Y%m")

    # ── Salvar snapshot ─────────────────────────────────────────────────────────
    base_dir = os.path.dirname(os.path.abspath(__file__))
    saida_consulta_dir = os.path.join(base_dir, "arquivo")
    os.makedirs(saida_consulta_dir, exist_ok=True)
    
    saida_consulta = os.path.join(saida_consulta_dir, f"processo_consulta_{ano_mes}.xlsx")
    df.to_excel(saida_consulta, index=False)
    print(f"Consulta salva: {saida_consulta}")

    # ── Classificação por tipo e extração de datas ──────────────────────────────
    df["Tipo"] = df[col_processo].astype(str).str.upper().apply(
        lambda x: "CAP" if "CAP" in x else "SEG"
    )
    df["Ano"] = df[col_data].dt.year
    df["Mês"] = df[col_data].dt.month

    # ── Histórico normalizado (tabela fato) ─────────────────────────────────────
    fato = (
        df.groupby([col_responsavel, "Tipo", "Ano", "Mês"])
        .size()
        .reset_index(name="Quantidade")
        .rename(columns={col_responsavel: "Responsável"})
    )

    historico_path = os.path.join(base_dir, "saida", "historico.xlsx")
    os.makedirs(os.path.dirname(historico_path), exist_ok=True)

    if os.path.exists(historico_path):
        try:
            historico = pd.read_excel(historico_path)
            # keep='last' garante que os dados recém-processados sobrescrevam os antigos em caso de choque
            fato = pd.concat([historico, fato]).drop_duplicates(
                subset=["Responsável", "Tipo", "Ano", "Mês"], 
                keep='last'
            )
        except Exception as e:
            print(f"Erro ao mesclar com o histórico existente: {e}")

    # Ordenar e salvar
    fato = fato.sort_values(["Ano", "Mês", "Responsável", "Tipo"]).reset_index(drop=True)
    fato.to_excel(historico_path, index=False)
    print(f"Histórico atualizado com sucesso: {historico_path}")


def main():
    arquivo_selecionado = selecionar_arquivo()
    
    if not arquivo_selecionado:
        print("Operação cancelada. Nenhum arquivo selecionado.")
        sys.exit(0)
        
    print(f"Arquivo selecionado: {arquivo_selecionado}\nIniciando processamento...")
    processar_dados_excel(arquivo_selecionado)


if __name__ == "__main__":
    main()