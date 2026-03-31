# ============================================
# LIMPEZA DE MÚLTIPLAS COLUNAS DE CIDADE
# Para 2 ou mais colunas no mesmo Excel
# ============================================

import pandas as pd
from rapidfuzz import process, fuzz
import re


# ============================================
# 1. CARREGAR DADOS
# ============================================

print("=" * 70)
print("CARREGANDO DADOS")
print("=" * 70)

# Ler o Excel
df = pd.read_excel("orders_full_flat_anonymized -tentativa.xlsx")
print(f"✅ Excel carregado: {len(df)} registros")

# Mostrar colunas disponíveis
print("\n📋 Colunas disponíveis no arquivo:")
for i, col in enumerate(df.columns):
    print(f"  {i}: {col}")

# Ler CSV de cidades
cidades = pd.read_csv(
    "C:/Users/Dataan/Downloads/Dados/Dados/worldcities.csv",
    encoding="utf-8"
)
lista_cidades = cidades["city_ascii"].dropna().astype(str).tolist()
print(f"\n✅ Lista de cidades carregada: {len(lista_cidades)} cidades únicas")


# ============================================
# 2. MANUAL CORRECTIONS DICTIONARY
# ============================================

manual_corrections = {
    # simples typos
    'londre': 'london',
    'londra': 'london',
    'londo': 'london',
    'londn': 'london',
    'lodnon': 'london',
    'londom': 'london',
    'londoan': 'london',
    'acton': 'london',
    'barking': 'london',
    'beckenham': 'london',
    'bermondsey': 'london',
    'bexleyheath': 'london',
    'bloomsbury': 'london',
    'bounds green': 'london',
    'brent': 'london',
    'brockley': 'london',
    'bromley': 'london',
    'camden': 'london',
    'caterham': 'london',
    'chelsea': 'london',
    'chislehurst': 'london',
    'chiswick': 'london',
    'clapham': 'london',
    'city of westminster': 'london',
    'croydon': 'london',
    'dulwich': 'london',
    'east ham': 'london',
    'edgware': 'london',
    'edgeware': 'london',
    'finchley': 'london',
    'gidea park': 'london',
    'greenford': 'london',
    'greenwich': 'london',
    'hammersmith and fulham': 'london',
    'hampstead': 'london',
    'hampstead way': 'london',
    'hanwell': 'london',
    'haringey': 'london',
    'harrow': 'london',
    'hayes': 'london',
    'hayes london': 'london',
    'hendon': 'london',
    'hillingdon': 'london',
    'hornchurch': 'london',
    'hounslow': 'london',
    'isleworth': 'london',
    'kingsbury': 'london',
    'kingston': 'london',
    'knightsbridge': 'london',
    'lambeth': 'london',
    'lewisham': 'london',
    'leytonstone': 'london',
    'leytonstone london': 'london',
    'mitcham': 'london',
    'morden': 'london',
    'new malden': 'london',
    'north kensington': 'london',
    'northwood': 'london',
    'orpington': 'london',
    'peckham': 'london',
    'queen elizabeth street': 'london',
    'rayners lane': 'london',
    'redbridge': 'london',
    'редбридж': 'london',
    'richmond upon thames': 'london',
    'romford': 'london',
    'ruislip': 'london',
    'ruslip': 'london',
    'south ruislip': 'london',
    'sidcup': 'london',
    'south croydon': 'london',
    'south grove': 'london',
    'southend': 'london',
    'stanmore': 'london',
    'stratford': 'london',
    'streatham vale': 'london',
    'sutton': 'london',
    'swanley': 'london',
    'tooting': 'london',
    'tottenham': 'london',
    'tower hamlets': 'london',
    'twickenham': 'london',
    'uxbridge': 'london',
    'walthamstow': 'london',
    'wembley': 'london',
    'wembley park': 'london',
    'wembley park london': 'london',
    'west dulwich': 'london',
    'west wickham': 'london',
    'wesr wickham': 'london',
    'white city estate': 'london',
    'wimbledon': 'london',
    'woodford green': 'london',
    'middlesex': 'london',
    'middlesex london': 'london',

    # abreviações
    'ny': 'new york',
    'nyc': 'new york',
    'new yorl': 'new york',
    
    # paris
    'paria': 'paris',
    'paris.': 'paris',
    
    # tokyo
    'toquio': 'tokyo',
    'tokio': 'tokyo',
    'toqui': 'tokyo',
    
    # endereços de londres (sempre em lowercase)
    'lo': 'london',
    'london england': 'london',
    'london e ej': 'london',
    'london e1 5ej': 'london',
    'london stanmore': 'london',
    'leytonstone london': 'london',
    'leytonstone': 'london',
    'hayes london': 'london',
    'middlesex london': 'london',
    'middlesex': 'london',
    
    # outros
    'church road, edgbaston': 'birmingham',
    'ewhurst': 'ewhurst',
    'birminham': 'birmingham',
    'queen elizabeth street': 'london',
    'warrington cheshire': 'warrington',
}



# ============================================
# 3. FUNÇÕES DE LIMPEZA
# ============================================

def fuzzy_match_city(city_name, valid_cities, threshold=80):
    """Combina usando fuzzy matching com múltiplos métodos"""
    city_clean = str(city_name).strip().lower()
    
    # Verificar exato primeiro
    if city_clean in [c.lower() for c in valid_cities]:
        return city_clean, 100, 'exact'
    
    # Tentar múltiplos métodos
    results = []
    
    match, score, _ = process.extractOne(
        city_clean, 
        valid_cities, 
        scorer=fuzz.token_sort_ratio
    )
    results.append((match, score, 'token_sort'))
    
    match, score, _ = process.extractOne(
        city_clean, 
        valid_cities, 
        scorer=fuzz.token_set_ratio
    )
    results.append((match, score, 'token_set'))
    
    match, score, _ = process.extractOne(
        city_clean, 
        valid_cities, 
        scorer=fuzz.ratio
    )
    results.append((match, score, 'ratio'))
    
    best = max(results, key=lambda x: x[1])
    
    if best[1] >= threshold:
        return best[0], best[1], best[2]
    else:
        return None, best[1], 'failed'


def clean_city_name(city_name, valid_cities, manual_corrections=manual_corrections):
    """Função principal com múltiplas estratégias"""
    
    if pd.isna(city_name) or str(city_name).strip() == '':
        return None, 0, 'empty'
    
    city_original = str(city_name).strip()
    city_clean = city_original.lower()
    
    # ESTRATÉGIA 1: Correções manuais
    result = manual_corrections.get(city_clean)
    if result:
        return result, 100, 'manual'
    
        # ESTRATÉGIA 2: Fuzzy matching (threshold menor) - APENAS se score >= 90
    match, score, method = fuzzy_match_city(city_clean, valid_cities, threshold=75)
    if match and score >= 90:
        return match, score, f'fuzzy_lenient_{method}'
    
    return None, 0, 'no_match'


# ============================================
# 4. DEFINIR COLUNAS A LIMPAR
# ============================================


CITY_COLUMNS = [
    'billingAddress.city',      
    'shippingAddress.city',     

]


print("\n" + "=" * 70)
print("VERIFICANDO COLUNAS")
print("=" * 70)

for col in CITY_COLUMNS:
    if col in df.columns:
        print(f"✅ Coluna encontrada: '{col}'")
    else:
        print(f"❌ Coluna NÃO encontrada: '{col}'")
        print(f"   Colunas disponíveis: {list(df.columns)}")


# ============================================
# 5. PROCESSAR TODAS AS COLUNAS
# ============================================

print("\n" + "=" * 70)
print("PROCESSANDO CIDADES")
print("=" * 70)

all_results = {}
statistics = {}

for col_name in CITY_COLUMNS:
    if col_name not in df.columns:
        print(f"\n⚠️  Pulando coluna '{col_name}' (não encontrada)")
        continue
    
    print(f"\n📍 Processando coluna: '{col_name}'")
    
    results = []
    
    for idx, original_value in enumerate(df[col_name], 1):
        
        if pd.isna(original_value) or str(original_value).strip() == "":
            results.append({
                'original': original_value,
                'corrected': None,
                'score': 0,
                'method': 'empty',
                'status': 'skipped',
                'confidence': 'N/A'
            })
            continue
        
        
        city_str = str(original_value).strip()
        if "," in city_str:
            city_str = str(city_str.split(",")[-1].strip())
        
        
        matched, score, method = clean_city_name(city_str, lista_cidades)
        
        
        if score >= 90:
            confidence = 'high'
            status = 'auto_corrected'
            final_city = matched
        elif score >= 80 and matched:
            confidence = 'medium'
            status = 'corrected'
            final_city = matched
        else:
            confidence = 'low'
            status = 'not_corrected'
            final_city = city_str if matched is None else matched
        
        results.append({
            'original': original_value,
            'extracted': city_str,
            'corrected': final_city,
            'score': score,
            'method': method,
            'status': status,
            'confidence': confidence
        })
        
        if idx % 500 == 0:
            print(f"  ✓ {idx} registros processados")
    
    results_df = pd.DataFrame(results)
    all_results[col_name] = results_df
    
    
    df[f'{col_name}_corrected'] = results_df['corrected']
    df[f'{col_name}_score'] = results_df['score']
    df[f'{col_name}_method'] = results_df['method']
    df[f'{col_name}_confidence'] = results_df['confidence']
    
    # Calcular estatísticas
    total = len(results_df)
    auto_corrected = (results_df['status'] == 'auto_corrected').sum()
    corrected = (results_df['status'] == 'corrected').sum()
    not_corrected = (results_df['status'] == 'not_corrected').sum()
    skipped = (results_df['status'] == 'skipped').sum()
    
    statistics[col_name] = {
        'total': total,
        'auto_corrected': auto_corrected,
        'corrected': corrected,
        'not_corrected': not_corrected,
        'skipped': skipped,
        'success_rate': (auto_corrected + corrected) / (total - skipped) * 100 if (total - skipped) > 0 else 0
    }
    
    print(f"  ✓ Coluna '{col_name}' processada: {len(results_df)} registros")


# ============================================
# 6. GERAR RELATÓRIO
# ============================================

print("\n" + "=" * 70)
print("GERANDO RELATÓRIO")
print("=" * 70)

for col_name, stats in statistics.items():
    print(f"\n📊 COLUNA: '{col_name}'")
    print(f"  Total processado: {stats['total']}")
    print(f"  ✓ Auto-corrigido (score ≥ 90): {stats['auto_corrected']}")
    print(f"  ✓ Corrigido (score ≥ 80): {stats['corrected']}")
    print(f"  ✗ Não corrigido: {stats['not_corrected']}")
    print(f"  ⊘ Vazio/Ignorado: {stats['skipped']}")
    print(f"  📈 Taxa de sucesso: {stats['success_rate']:.1f}%")


# ============================================
# 7. CASOS PARA REVISAR
# ============================================

print("\n" + "=" * 70)
print("⚠️  CASOS PARA REVISAR (score 75-85)")
print("=" * 70)

for col_name, results_df in all_results.items():
    borderline = results_df[(results_df['score'] >= 75) & (results_df['score'] < 90)]
    
    if len(borderline) > 0:
        print(f"\n📍 Coluna '{col_name}': {len(borderline)} casos")
        for idx, row in borderline.head(10).iterrows():
            print(f"  '{row['extracted']}' → '{row['corrected']}' (score: {row['score']})")
        if len(borderline) > 10:
            print(f"  ... e mais {len(borderline) - 10}")
    else:
        print(f"\n📍 Coluna '{col_name}': ✅ Nenhum caso borderline")


# ============================================
# 8. TOP CIDADES POR COLUNA
# ============================================

print("\n" + "=" * 70)
print("📊 TOP 10 CIDADES POR COLUNA")
print("=" * 70)

for col_name in CITY_COLUMNS:
    if col_name not in df.columns:
        continue
    
    corrected_col = f'{col_name}_corrected'
    print(f"\n📍 {col_name}:")
    top_cities = df[corrected_col].value_counts().head(10)
    for city, count in top_cities.items():
        print(f"  {city}: {count} registros")


# ============================================
# 9. SALVAR RESULTADOS
# ============================================

print("\n" + "=" * 70)
print("SALVANDO ARQUIVOS")
print("=" * 70)

# Arquivo principal
df.to_excel("meuarquivo_corrigido.xlsx", index=False)
print(f"✅ Arquivo principal: meuarquivo_corrigido.xlsx")

# Auditoria para cada coluna
for col_name, results_df in all_results.items():
    safe_col_name = col_name.replace(".", "_").replace("/", "_")
    audit_path = f"auditoria_{safe_col_name}.csv"
    results_df.to_csv(audit_path, index=False, encoding='utf-8')
    print(f"✅ Auditoria: {audit_path}")

# Casos para revisar de cada coluna
for col_name, results_df in all_results.items():
    problematic = results_df[results_df['status'].isin(['not_corrected', 'skipped'])]
    if len(problematic) > 0:
        safe_col_name = col_name.replace(".", "_").replace("/", "_")
        review_path = f"revisar_{safe_col_name}.csv"
        problematic.to_csv(review_path, index=False, encoding='utf-8')
        print(f"✅ Para revisar: {review_path} ({len(problematic)} registros)")

# Sumário
with open("sumario_limpeza.txt", "w", encoding='utf-8') as f:
    f.write("=" * 70 + "\n")
    f.write("RESUMO DA LIMPEZA DE MÚLTIPLAS COLUNAS DE CIDADES\n")
    f.write("=" * 70 + "\n\n")
    
    for col_name, stats in statistics.items():
        f.write(f"COLUNA: {col_name}\n")
        f.write(f"  Total processado: {stats['total']}\n")
        f.write(f"  Auto-corrigido: {stats['auto_corrected']}\n")
        f.write(f"  Corrigido: {stats['corrected']}\n")
        f.write(f"  Não corrigido: {stats['not_corrected']}\n")
        f.write(f"  Vazio: {stats['skipped']}\n")
        f.write(f"  Taxa de sucesso: {stats['success_rate']:.1f}%\n\n")

print(f"✅ Sumário: sumario_limpeza.txt")


# ============================================
# 10. CONCLUSÃO
# ============================================

print("\n" + "=" * 70)
print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
print("=" * 70)
print("\nArquivos gerados:")
print("  • meuarquivo_corrigido.xlsx - Dados com todas as correções")
print("  • auditoria_*.csv - Detalhes de cada coluna processada")
print("  • revisar_*.csv - Casos que precisam de revisão manual")
print("  • sumario_limpeza.txt - Resumo das estatísticas")
print("\n✨ Próximos passos:")
print("  1. Verifique 'revisar_*.csv' para casos problemáticos")
print("  2. Atualize MANUAL_CORRECTIONS com padrões encontrados")
print("  3. Re-execute para melhorar os resultados")
