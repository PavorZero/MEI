import unicodedata
import pandas as pd
import re

def normalize_name(name):
    """
    Normaliza o texto para melhorar a comparação:
    - Remove acentos, mantendo letras (é -> e, ã -> a).
    - Converte para minúsculas.
    - Remove caracteres não alfabéticos.
    - Remove espaços extras, tabs e quebras de linha.
    """
    # Substitui quebras de linha e tabs por espaço
    name = name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    
    # Remove acentuação, mantendo os caracteres base (ex: é -> e)
    name = unicodedata.normalize('NFKD', name)
    name = ''.join(c for c in name if not unicodedata.combining(c))
    
    # Converte para minúsculas
    name = name.lower()

    # Remove múltiplos espaços e trim
    name = re.sub(r'\s+', ' ', name).strip()
    
    # Remove qualquer caractere que não seja letra ou espaço (opcional)
    # name = re.sub(r'[^a-z\s]', '', name)

    return name

def processar_lista_nomes(arquivo_entrada, arquivo_saida):
    """
    Lê uma lista de nomes de um arquivo, normaliza e salva em Excel.
    """
    try:
        # Lê os nomes do arquivo (assumindo um nome por linha)
        with open(arquivo_entrada, 'r', encoding='utf-8') as f:
            nomes = [linha.strip() for linha in f.readlines()]
        
        # Normaliza cada nome
        nomes_normalizados = [normalize_name(nome) for nome in nomes]
        
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
