#criado por PavorZero
import pandas as pd
import os

def clean_cell_content(cell):
    """Remove quebras de linha, tabs e espaços duplicados."""
    if pd.isna(cell):
        return cell
    return str(cell).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()

def clean_columns_in_sheet(input_file, sheet_name, column_name_or_all="*", output_file=None):
    """
    Corrige quebras de linha em uma ou todas as colunas dentro de uma aba do Excel.
    
    Args:
        input_file (str): Caminho para o arquivo Excel.
        sheet_name (str): Nome da aba a ser processada.
        column_name_or_all (str): Nome da coluna a corrigir, ou '*' para todas.
        output_file (str, opcional): Caminho de saída. Se não fornecido, salva como '<arquivo>_corrigido.xlsx'.
    """
    # Carrega todas as abas
    excel_data = pd.read_excel(input_file, sheet_name=None, dtype=str)

    # Verifica se a aba existe
    if sheet_name not in excel_data:
        raise ValueError(f"A aba '{sheet_name}' não foi encontrada. Abas disponíveis: {list(excel_data.keys())}")

    df = excel_data[sheet_name]

    # Decide se corrige uma ou todas as colunas
    if column_name_or_all == "*":
        print(f"Corrigindo todas as colunas da aba '{sheet_name}'...")
        for col in df.columns:
            df[col] = df[col].apply(clean_cell_content)
    else:
        if column_name_or_all not in df.columns:
            raise ValueError(f"A coluna '{column_name_or_all}' não foi encontrada na aba '{sheet_name}'. Colunas disponíveis: {list(df.columns)}")
        print(f"Corrigindo coluna '{column_name_or_all}' da aba '{sheet_name}'...")
        df[column_name_or_all] = df[column_name_or_all].apply(clean_cell_content)

    # Atualiza a aba no dicionário
    excel_data[sheet_name] = df

    # Define caminho de saída
    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_corrigido{ext}"

    # Salva todas as abas, mantendo as não modificadas
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for name, data in excel_data.items():
            data.to_excel(writer, sheet_name=name, index=False)

    print(f"Aba '{sheet_name}' corrigida com sucesso.")
    print(f"Arquivo salvo como: {output_file}")


# ========================
# CONFIGURAÇÕES MANUAIS
# ========================
if __name__ == "__main__":
    # Caminho do arquivo Excel
    input_excel_path = r"C:\Users\joao.silva\Documents\RELATORIOS MEI\RELATÓRIO GERAL.xlsx"

    # Nome da aba a ser corrigida
    aba_para_corrigir = "2020"

    # Corrigir apenas uma coluna (ex: 'names_d') ou todas as colunas ('*')
    coluna_para_corrigir = "*"  # Exemplo: 'names_d' ou '*' para todas

    # (Opcional) Nome do arquivo de saída
    output_excel_path = None  # Ou defina o caminho completo

    clean_columns_in_sheet(input_excel_path, aba_para_corrigir, coluna_para_corrigir, output_excel_path)
