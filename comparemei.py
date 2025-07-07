#criado por PavorZero
from datetime import datetime
from rapidfuzz import fuzz, process
import unicodedata
import re
import pandas as pd

def normalize_name(name):
    """
    Normaliza o texto para melhorar a comparação:
    - Remove acentos e caracteres especiais.
    - Converte para letras minúsculas.
    - Remove espaços extras e quebras de linha.
    - Remove caracteres não alfabéticos (opcional, dependendo do caso de uso).
    
    Args:
        name (str): Nome a ser normalizado.
    
    Returns:
        str: Nome normalizado.
    """
    # Remove quebras de linha primeiro
    name = name.replace('\n', ' ').replace('\r', '')
    
    # Remove acentos e caracteres especiais
    name = unicodedata.normalize('NFKD', name).encode('ASCII', 'ignore').decode('utf-8')
    # Converte para letras minúsculas
    name = name.lower()
    # Remove espaços extras e normaliza espaços
    name = re.sub(r'\s+', ' ', name).strip()
    # Remove caracteres não alfabéticos (opcional)
    name = re.sub(r'[^a-z\s]', '', name)
    return name

def compare_similar_names(file_path_a, file_path_b, similarity_threshold=85):
    """
    Compara nomes de dois arquivos de texto, identificando nomes semelhantes com base em uma pontuação de similaridade.
    
    Args:
        file_path_a (str): Caminho para o primeiro arquivo de nomes.
        file_path_b (str): Caminho para o segundo arquivo de nomes.
        similarity_threshold (int): Pontuação mínima de similaridade (0-100) para considerar dois nomes como iguais.
    
    Returns:
        list: Lista de pares de nomes semelhantes.
    """
    print("Início da comparação:", datetime.now())
    
    # Lendo os arquivos com codificação UTF-8 e removendo quebras de linha
    with open(file_path_a, 'r', encoding='utf-8') as file_a:
        names_a = [line.strip() for line in file_a if line.strip()]
    
    with open(file_path_b, 'r', encoding='utf-8') as file_b:
        names_b = [line.strip() for line in file_b if line.strip()]
    
    # Normalizando os nomes
    names_a = [normalize_name(name) for name in names_a]
    names_b = [normalize_name(name) for name in names_b]
    
    # Lista para armazenar os pares de nomes semelhantes
    similar_names = []
    
    # Comparando cada nome da lista A com os nomes da lista B
    for i, name_a in enumerate(names_a):
        # Encontrar o nome mais semelhante na lista B
        match, score, _ = process.extractOne(name_a, names_b, scorer=fuzz.ratio)
        
        # Verificar se a similaridade está acima do limiar
        if score >= similarity_threshold:
            similar_names.append((name_a, match, score))
        
        # Exibir progresso a cada 100 iterações
        if (i + 1) % 100 == 0:
            print(f"Progresso: {i + 1}/{len(names_a)} nomes processados - {datetime.now()}")
    
    print("Quantidade de nomes semelhantes encontrados:", len(similar_names))
    print("Fim da comparação:", datetime.now())
    
    return similar_names

def save_similar_names_to_excel(similar_names, output_file):
    """
    Salva os pares de nomes semelhantes em um arquivo Excel.
    
    Args:
        similar_names (list): Lista de pares de nomes semelhantes.
        output_file (str): Caminho para o arquivo de saída.
    """
    # Criando um DataFrame com os dados
    df = pd.DataFrame(similar_names, columns=['Nome Lista A', 'Nome Lista B', 'Similaridade (%)'])
    
    # Configurando o pandas para não quebrar linhas
    pd.set_option('display.max_colwidth', None)
    
    # Salvando o DataFrame em um arquivo Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    print(f"Nomes semelhantes salvos em '{output_file}'")

if __name__ == "__main__":
    # Caminhos para os arquivos de texto
    file_path_a = 'names_a.txt'
    file_path_b = 'names_b.txt'
    
    # Comparando os nomes nos arquivos
    similar_names = compare_similar_names(file_path_a, file_path_b, similarity_threshold=85)
    
    # Salvando os resultados em um arquivo Excel
    save_similar_names_to_excel(similar_names, 'similar_names.xlsx')
