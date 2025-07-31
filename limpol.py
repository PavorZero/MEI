import pandas as pd
import os

def clean_cell_content(cell):
    """Remove quebras de linha, tabs e espaços duplicados."""
    if pd.isna(cell):
        return cell
    return str(cell).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()

def clean_columns_in_sheet(input_file, sheet_name="*", column_name_or_all="*", output_file=None):
    """
    Corrige quebras de linha em uma ou todas as colunas de uma ou todas as abas do Excel.

    Args:
        input_file (str): Caminho para o arquivo Excel.
        sheet_name (str): Nome da aba a ser processada, ou '*' para todas.
        column_name_or_all (str): Nome da coluna a corrigir, ou '*' para todas.
        output_file (str, opcional): Caminho de saída. Se não fornecido, salva como '<arquivo>_corrigido.xlsx'.
    """
    # Carrega todas as abas
    excel_data = pd.read_excel(input_file, sheet_name=None, dtype=str)

    sheets_to_process = list(excel_data.keys()) if sheet_name == "*" else [sheet_name]

    # Verifica se as abas existem
    for s in sheets_to_process:
        if s not in excel_data:
            raise ValueError(f"A aba '{s}' não foi encontrada. Abas disponíveis: {list(excel_data.keys())}")

    for s in sheets_to_process:
        df = excel_data[s]
        print(f"Processando aba '{s}'...")

        if column_name_or_all == "*":
            print("  Corrigindo todas as colunas...")
            for col in df.columns:
                df[col] = df[col].apply(clean_cell_content)
        else:
            if column_name_or_all not in df.columns:
                raise ValueError(f"A coluna '{column_name_or_all}' não foi encontrada na aba '{s}'. Colunas disponíveis: {list(df.columns)}")
            print(f"  Corrigindo coluna '{column_name_or_all}'...")
            df[column_name_or_all] = df[column_name_or_all].apply(clean_cell_content)

        # Atualiza a aba modificada
        excel_data[s] = df

    # Define caminho de saída
    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_corrigido{ext}"

    # Salva todas as abas no novo arquivo
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for name, data in excel_data.items():
            data.to_excel(writer, sheet_name=name, index=False)

    print("Limpeza concluída.")
    print(f"Arquivo salvo como: {output_file}")

# ========================
# CONFIGURAÇÕES MANUAIS
# ========================
if __name__ == "__main__":
    input_excel_path = r"C:\Users\joao.silva\Documents\RELATORIOS MEI\Contato-DTI_20250701.xlsx"
    aba_para_corrigir = "*"          # Ex: "2019" ou "*" para todas
    coluna_para_corrigir = "*"      # Ex: "nome_coluna" ou "*" para todas
    output_excel_path = None        # Ou defina um caminho como "saida.xlsx"

    clean_columns_in_sheet(input_excel_path, aba_para_corrigir, coluna_para_corrigir, output_excel_path)
