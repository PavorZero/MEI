import pandas as pd

# Lista de sufixos atualizada
sufixos_femininos = ['a', 'e', 'ia', 'ice', 'na', 'ra', 'da', 'ta', 'sa', 'ca', 'la', 
                    'ara', 'ana', 'ina', 'ela', 'isa', 'lia', 'nia', 'ria', 'ta', 'na']
sufixos_masculinos = ['o', 'ão', 'io', 'os', 'us', 'as', 'es', 'to', 'do', 'ro', 
                     'ardo', 'berto', 'cio', 'dio', 'elio', 'gio', 'lio', 'nio', 'rdo']

def extrair_primeiro_nome(nome_completo):
    """Extrai o primeiro nome de um nome completo"""
    return nome_completo.strip().split()[0]

def identificar_genero(nome):
    """
    Tenta identificar o gênero com base no nome.
    Retorna 'Feminino', 'Masculino' ou 'Indeterminado'
    """
    nome = nome.strip().lower()
    
    # Verifica sufixos femininos
    for sufixo in sufixos_femininos:
        if nome.endswith(sufixo):
            return 'Feminino'
    
    # Verifica sufixos masculinos
    for sufixo in sufixos_masculinos:
        if nome.endswith(sufixo):
            return 'Masculino'
    
    return 'Indeterminado'

def processar_arquivo_txt(caminho_arquivo):
    """Lê o arquivo TXT e retorna lista de nomes completos"""
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as file:
            return [linha.strip() for linha in file.readlines() if linha.strip()]
    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_arquivo}' não encontrado.")
        return []

# Processamento principal
nomes_completos = processar_arquivo_txt('names_c.txt')

if not nomes_completos:
    print("Usando lista de exemplo pois não foram encontrados nomes no arquivo.")
    nomes_completos = [
        'Maria Silva', 
        'João Pedro Santos', 
        'Ana Paula Oliveira', 
        'José Pereira', 
        'Carlos Alberto Souza'
    ]

# Processa os nomes
resultados = []
for nome_completo in nomes_completos:
    primeiro_nome = extrair_primeiro_nome(nome_completo)
    genero = identificar_genero(primeiro_nome)
    resultados.append({
        'Nome Completo': nome_completo,
        'Primeiro Nome': primeiro_nome,
        'Gênero': genero
    })

# Cria DataFrame
df = pd.DataFrame(resultados)

# Exporta para Excel
arquivo_excel = 'genero_por_nome.xlsx'
df.to_excel(arquivo_excel, index=False, engine='openpyxl')

print(f"\nResumo:")
print(f"Total de nomes processados: {len(nomes_completos)}")
print(f"Feminino: {len(df[df['Gênero'] == 'Feminino'])}")
print(f"Masculino: {len(df[df['Gênero'] == 'Masculino'])}")
print(f"Indeterminado: {len(df[df['Gênero'] == 'Indeterminado'])}")
print(f"\nResultados exportados para '{arquivo_excel}'")