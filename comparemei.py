from datetime import datetime
from rapidfuzz import fuzz, process
import unicodedata
import re
import pandas as pd
from collections import defaultdict
from itertools import combinations
import string
from openpyxl.utils import get_column_letter

def normalize_text(name):
    """Normalização completa do texto antes da comparação"""
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

def extract_first_last_name(name):
    """Extrai primeiro e último nome de uma string normalizada"""
    parts = name.split()
    if not parts:
        return '', ''
    first = parts[0]
    last = parts[-1] if len(parts) > 1 else ''
    return first, last

def clean_final_string(s):
    return str(s).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()

def load_names(file_paths):
    names_dict = {}
    for key, path in file_paths.items():
        with open(path, 'r', encoding='utf-8') as file:
            lines = [line.strip() for line in file if line.strip()]
            names_dict[key] = {
                'original': lines,
                'normalized': [normalize_text(name) for name in lines],
                'first_last': [extract_first_last_name(normalize_text(name)) for name in lines]
            }
    return names_dict

def compare_names(name1, name2, similarity_threshold):
    """Compara nomes usando primeiro+último nome OU similaridade fuzzy"""
    first1, last1 = name1
    first2, last2 = name2
    
    # Se primeiro E último nome forem iguais, considera match perfeito
    if first1 == first2 and last1 == last2 and first1 and last1:
        return 100  # Match perfeito
    
    # Caso contrário, usa similaridade fuzzy
    full_name1 = f"{first1} {last1}" if last1 else first1
    full_name2 = f"{first2} {last2}" if last2 else first2
    return fuzz.ratio(full_name1, full_name2)

def find_flexible_matches(names_dict, selected_lists, similarity_threshold=85, min_matches=2):
    print(f"\nIniciando comparação... {datetime.now()}")
    
    results = defaultdict(lambda: {'lists': [], 'matches': {}, 'scores': {}})
    total_comparisons = len(list(combinations(selected_lists, 2)))
    current_comparison = 0
    
    for list1, list2 in combinations(selected_lists, 2):
        current_comparison += 1
        print(f"\nComparando {list1} ↔ {list2} ({current_comparison}/{total_comparisons})...")
        
        names1 = names_dict[list1]['first_last']
        orig1 = names_dict[list1]['original']
        names2 = names_dict[list2]['first_last']
        orig2 = names_dict[list2]['original']
        
        for i, name1 in enumerate(names1):
            best_score = 0
            best_match_idx = -1
            
            for j, name2 in enumerate(names2):
                score = compare_names(name1, name2, similarity_threshold)
                if score > best_score:
                    best_score = score
                    best_match_idx = j
            
            if best_score >= similarity_threshold:
                original_name1 = orig1[i]
                original_name2 = orig2[best_match_idx]
                
                for norm_name, lst, orig in [
                    (f"{name1[0]} {name1[1]}", list1, original_name1),
                    (f"{names2[best_match_idx][0]} {names2[best_match_idx][1]}", list2, original_name2)
                ]:
                    if lst not in results[norm_name]['lists']:
                        results[norm_name]['lists'].append(lst)
                        results[norm_name]['matches'][lst] = orig
                        results[norm_name]['scores'][lst] = best_score

    final_results = {}
    for norm_name, data in results.items():
        if len(data['lists']) >= min_matches:
            base_name = next(iter(data['matches'].values()))
            final_results[base_name] = {
                'lists': data['lists'],
                'matches': data['matches'],
                'scores': data['scores'],
                'normalized': norm_name
            }

    print(f"\nProcesso finalizado. {len(final_results)} matches encontrados.")
    return final_results

def save_results_to_excel(results, output_file, selected_lists):
    if not results:
        print("Nenhum resultado para salvar.")
        return

    data = []
    for base_name, info in results.items():
        row = {
            'Nome Base': base_name,
            'Listas': ', '.join(info['lists']),
            'Normalizado': info['normalized']
        }
        for lst in selected_lists:
            if lst in info['matches']:
                row[f'Nome {lst}'] = info['matches'][lst]
                row[f'Similaridade {lst}'] = info['scores'].get(lst, 'N/A')
            else:
                row[f'Nome {lst}'] = 'N/A'
                row[f'Similaridade {lst}'] = 'N/A'
        data.append(row)
    
    df = pd.DataFrame(data)
    columns = ['Nome Base', 'Normalizado', 'Listas']
    for lst in selected_lists:
        columns.extend([f'Nome {lst}', f'Similaridade {lst}'])
    
    df = df[columns]

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultados')
        worksheet = writer.sheets['Resultados']
        
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
    max_matches = len(selected_lists)
    
    similarity = int(input(f"\nLimiar de similaridade (1-100) [85]: ") or 85)
    min_matches = int(input(f"Matches mínimos (2-{max_matches}) [2]: ") or 2)
    
    names_data = load_names(available_files)
    matches = find_flexible_matches(
        names_data,
        selected_lists,
        similarity_threshold=similarity,
        min_matches=min_matches
    )
    
    output_file = input("\nNome do arquivo de saída [resultados.xlsx]: ") or "resultados.xlsx"
    save_results_to_excel(matches, output_file, selected_lists)

if __name__ == "__main__":
    main()
