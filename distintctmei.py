#criado por PavorZero
from datetime import datetime
import unicodedata
import re
import pandas as pd

def normalize_name(name):
    """
    Normaliza o texto para melhorar a comparação:
    - Remove acentos e caracteres especiais.
    - Converte para letras minúsculas.
    - Remove espaços extras.
    - Remove caracteres não alfabéticos (opcional, dependendo do caso de uso).
    
    Args:
        name (str): Nome a ser normalizado.
    
    Returns:
        str: Nome normalizado.
    """
    # Remove acentos e caracteres especiais
    name = unicodedata.normalize('NFKD', name).encode('ASCII', 'ignore').decode('utf-8')
    # Converte para letras minúsculas
    name = name.lower()
    # Remove espaços extras
    name = re.sub(r'\s+', ' ', name).strip()
    # Remove caracteres não alfabéticos (opcional)
    name = re.sub(r'[^a-z\s]', '', name)
    return name

def find_distinct_names(file_path_a, file_path_b):
    """
    Identifica nomes distintos entre duas listas.
    
    Args:
        file_path_a (str): Caminho para o primeiro arquivo de nomes.
        file_path_b (str): Caminho para o segundo arquivo de nomes.
    
    Returns:
        tuple: Dois conjuntos contendo os nomes distintos de cada lista.
    """
    print("Início da análise:", datetime.now())
    
    # Lendo os arquivos com codificação UTF-8
    with open(file_path_a, 'r', encoding='utf-8') as file_a:
        lista_a = file_a.read().splitlines()
    
    with open(file_path_b, 'r', encoding='utf-8') as file_b:
        lista_b = file_b.read().splitlines()
    
    # Normalizando os nomes
    lista_a = [normalize_name(name) for name in lista_a]
    lista_b = [normalize_name(name) for name in lista_b]
    
    # Convertendo para conjuntos para facilitar a comparação
    conjunto_a = set(lista_a)
    conjunto_b = set(lista_b)
    
    # Encontrando nomes distintos
    distintos_a = conjunto_a - conjunto_b  # Nomes que estão na lista A, mas não na lista B
    distintos_b = conjunto_b - conjunto_a  # Nomes que estão na lista B, mas não na lista A
    
    print("Quantidade de nomes distintos na Lista A:", len(distintos_a))
    print("Quantidade de nomes distintos na Lista B:", len(distintos_b))
    print("Fim da análise:", datetime.now())
    
    return distintos_a, distintos_b

def save_distinct_names_to_excel(distintos_a, distintos_b, output_file):
    """
    Salva os nomes distintos em um arquivo Excel.
    
    Args:
        distintos_a (set): Nomes distintos da Lista A.
        distintos_b (set): Nomes distintos da Lista B.
        output_file (str): Caminho para o arquivo de saída.
    """
    # Convertendo os conjuntos para listas
    lista_a = list(distintos_a)
    lista_b = list(distintos_b)
    
    # Garantindo que ambas as listas tenham o mesmo tamanho
    max_length = max(len(lista_a), len(lista_b))
    lista_a.extend([None] * (max_length - len(lista_a)))  # Preenche com None
    lista_b.extend([None] * (max_length - len(lista_b)))  # Preenche com None
    
    # Criando um DataFrame com os dados
    df = pd.DataFrame({
        'Nomes Distintos Lista A': lista_a,
        'Nomes Distintos Lista B': lista_b
    })
    
    # Salvando o DataFrame em um arquivo Excel
    df.to_excel(output_file, index=False)
    print(f"Nomes distintos salvos em '{output_file}'")

if __name__ == "__main__":
    # Caminhos para os arquivos de texto
    file_path_a = 'names_a.txt'
    file_path_b = 'names_b.txt'
    
    # Encontrando os nomes distintos
    distintos_a, distintos_b = find_distinct_names(file_path_a, file_path_b)
    
    # Salvando os resultados em um arquivo Excel
    save_distinct_names_to_excel(distintos_a, distintos_b, 'distinct_names.xlsx')
