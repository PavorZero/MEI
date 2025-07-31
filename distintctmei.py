from datetime import datetime
from rapidfuzz import fuzz
import unicodedata
import re
import pandas as pd
from collections import defaultdict
from itertools import combinations
import string
from openpyxl.utils import get_column_letter

def normalize_text(name):
    if not isinstance(name, str):
        name = str(name)
    name = re.sub(r'[\n\r\t]', ' ', name)
    name = unicodedata.normalize('NFKD', name)
    name = ''.join(c for c in name if not unicodedata.combining(c))
    name = name.lower()
    name = re.sub(r'[^a-z0-9\s]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    name = name.translate(str.maketrans('', '', string.punctuation))
    return name

def split_first_last(name):
    parts = name.strip().split()
    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], parts[0]
    return parts[0], parts[-1]

def load_names(file_paths):
    names_dict = {}
    for key, path in file_paths.items():
        with open(path, 'r', encoding='utf-8') as file:
            lines = [line.strip() for line in file if line.strip()]
            names_dict[key] = {
                'original': lines,
                'normalized': [normalize_text(name) for name in lines]
            }
    return names_dict

def find_differences(names_dict, selected_lists, similarity_threshold=85):
    print(f"\nIniciando comparação... {datetime.now()}")
    
    differences = []

    for list1, list2 in combinations(selected_lists, 2):
        print(f"\nComparando {list1} ↔ {list2}...")
        
        names1 = names_dict[list1]['normalized']
        orig1 = names_dict[list1]['original']
        names2 = names_dict[list2]['normalized']
        orig2 = names_dict[list2]['original']
        
        for i, name1 in enumerate(names1):
            first1, last1 = split_first_last(name1)
            best_match = None
            best_score = 0
            matched_name2 = None

            for j, name2 in enumerate(names2):
                first2, last2 = split_first_last(name2)
                score_first = fuzz.ratio(first1, first2)
                score_last = fuzz.ratio(last1, last2)
                avg_score = (score_first + score_last) / 2

                if first1 == first2 and last1 == last2:
                    best_match = None
                    best_score = 100
                    break  # nomes considerados iguais

                if avg_score > best_score:
                    best_score = avg_score
                    matched_name2 = orig2[j]
            
            if best_score < 100:
                differences.append({
                    'Lista 1': list1,
                    'Nome 1': orig1[i],
                    'Lista 2': list2,
                    'Nome 2': matched_name2 if matched_name2 else "Nenhuma correspondência",
                    'Similaridade (%)': round(best_score, 2)
                })

    print(f"\nProcesso finalizado. {len(differences)} nomes distintos identificados.")
    return differences

def save_differences_to_excel(results, output_file):
    if not results:
        print("Nenhum nome distinto encontrado.")
        return

    df = pd.DataFrame(results)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Nomes Distintos')
        worksheet = writer.sheets['Nomes Distintos']
        
        for col in df.columns:
            idx = df.columns.get_loc(col)
            letter = get_column_letter(idx + 1)
            max_length = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[letter].width = max_length

    print(f"Arquivo salvo: {output_file}")

def main():
    available_files = {
        'A': 'names_a.txt',
        'B': 'names_b.txt',
        'C': 'names_c.txt',
        'D': 'names_D.txt',
        'E': 'names_E.txt',
        'F': 'names_F.txt'
    }
    
    print("Arquivos disponíveis:")
    for k, v in available_files.items():
        print(f"{k}: {v}")

    selected = input("\nSelecione as listas (ex: AB, ABC): ").upper()
    while len(selected) < 2 or not all(c in available_files for c in selected):
        selected = input("Selecione pelo menos 2 listas válidas: ").upper()
    
    selected_lists = list(selected)

    similarity = int(input(f"\nLimiar de similaridade (1-100) [85]: ") or 85)
    output_file = input("\nNome do arquivo de saída [nomes_distintos.xlsx]: ") or "nomes_distintos.xlsx"
    
    names_data = load_names(available_files)
    differences = find_differences(names_data, selected_lists, similarity_threshold=similarity)
    save_differences_to_excel(differences, output_file)

if __name__ == "__main__":
    main()
