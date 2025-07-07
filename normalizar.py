import unicodedata
import pandas as pd

def normalizar_nome(nome):
    """
    Normaliza um nome removendo acentos, caracteres especiais e convertendo para minúsculas.
    """
    # Normaliza para a forma NFKD que separa caracteres e seus acentos
    nome_normalizado = unicodedata.normalize('NFKD', nome)
    # Remove os caracteres de acentuação e outros diacríticos
    nome_sem_acentos = ''.join([c for c in nome_normalizado if not unicodedata.combining(c)])
    # Substitui ç por c
    nome_sem_cedilha = nome_sem_acentos.replace('ç', 'c').replace('Ç', 'c')
    # Converte para minúsculas
    nome_minusculo = nome_sem_cedilha.lower()
    return nome_minusculo

def processar_lista_nomes(arquivo_entrada, arquivo_saida):
    """
    Lê uma lista de nomes de um arquivo, normaliza e salva em Excel.
    """
    try:
        # Lê os nomes do arquivo (assumindo um nome por linha)
        with open(arquivo_entrada, 'r', encoding='utf-8') as f:
            nomes = [linha.strip() for linha in f.readlines()]
        
        # Normaliza cada nome
        nomes_normalizados = [normalizar_nome(nome) for nome in nomes]
        
        # Cria um DataFrame pandas
        df = pd.DataFrame({
            'Nome Original': nomes,
            'Nome Normalizado': nomes_normalizados
        })
        
        # Salva como arquivo Excel
        df.to_excel(arquivo_saida, index=False)
        
        print(f"Processamento concluído! Arquivo salvo como '{arquivo_saida}'")
        
    except FileNotFoundError:
        print(f"Erro: Arquivo '{arquivo_entrada}' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

# Exemplo de uso
if __name__ == "__main__":
    arquivo_entrada = 'names_c.txt'  # Arquivo de entrada com os nomes
    arquivo_saida = 'nomes_normalizados.xlsx'  # Arquivo Excel de saída
    
    processar_lista_nomes(arquivo_entrada, arquivo_saida)