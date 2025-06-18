# Comparador de Nomes
Este repositório contém dois scripts em Python para comparar listas de nomes. Os scripts permitem identificar nomes semelhantes e nomes distintos entre duas listas fornecidas em arquivos de texto. Os resultados são exportados para um arquivo Excel para facilitar a análise.

Funcionalidades
# 1. Comparar Nomes Semelhantes
Identifica nomes que são semelhantes entre duas listas, mesmo que estejam escritos de forma ligeiramente diferente.
Usa a biblioteca rapidfuzz para calcular a similaridade entre os nomes com base na distância de Levenshtein.
Exporta os resultados para um arquivo Excel com as colunas:
- Nome Lista A
- Nome Lista B
- Similaridade (%)

# - 2. Identificar Nomes Distintos
Identifica nomes que estão presentes em uma lista, mas não na outra.
Exporta os resultados para um arquivo Excel com as colunas:
- Nomes Distintos Lista A
- Nomes Distintos Lista B

# - Requisitos
Certifique-se de que você tenha o Python 3.6 ou superior instalado. Além disso, instale as dependências necessárias:
 - pip install pandas rapidfuzz

# Como Usar
1. Comparar Nomes Semelhantes
Arquivos de Entrada
Crie dois arquivos de texto (names_a.txt e names_b.txt), onde cada linha contém um nome.

# Exemplo de Arquivos
- names_a.txt:
Pedro P Paulo
Maria Clara
João Silva
Ana Beatriz

- names_b.txt:
Pedro Phelipe Paulo
Maria C.
João da Silva
Ana B. Beatriz

# - Executar o Script
Use o seguinte comando para executar o script de comparação de nomes semelhantes:
python compare_similar_names.py

# - Resultado
O script gera um arquivo Excel chamado similar_names.xlsx com os nomes semelhantes e suas respectivas pontuações de similaridade.

# 2. Identificar Nomes Distintos
Arquivos de Entrada
Crie dois arquivos de texto (names_a.txt e names_b.txt), onde cada linha contém um nome.

# - Exemplo de Arquivos
- names_a.txt:
Pedro P Paulo
Maria Clara
João Silva
Ana Beatriz

- names_b.txt:
Pedro Phelipe Paulo
Maria C.
João da Silva
Ana B. Beatriz

# - Executar o Script
Use o seguinte comando para executar o script de identificação de nomes distintos:
python find_distinct_names.py

# Resultado
O script gera um arquivo Excel chamado distinct_names.xlsx com os nomes que estão presentes em uma lista, mas não na outra.

# - Personalização
Ajustar o Limiar de Similaridade
No script compare_similar_names.py, você pode ajustar o limiar de similaridade alterando o valor do parâmetro similarity_threshold. Por exemplo:
similar_names = compare_similar_names(file_path_a, file_path_b, similarity_threshold=90)
- Valores mais altos (ex.: 90): Mais rigoroso, encontra apenas nomes muito semelhantes.
- Valores mais baixos (ex.: 75): Mais permissivo, encontra nomes com menor grau de similaridade.

# - Dependências
As bibliotecas necessárias são:
- pandas: Para manipulação de dados e exportação para Excel.
- rapidfuzz: Para cálculo de similaridade entre strings.

# Instale as dependências com:
pip install pandas rapidfuzz

# - Contribuição
Sinta-se à vontade para abrir issues ou enviar pull requests para melhorias no código ou na documentação.
