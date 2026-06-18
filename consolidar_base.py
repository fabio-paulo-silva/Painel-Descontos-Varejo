"""
consolidar_base.py — versão vetorizada
Processa Descontos RGB com a mesma lógica do processar_desconto_boticario.py,
mas usando operações vetorizadas (sem .apply por grupo).
Merge com Cupons apenas para Vendedor e Canal de venda.
"""
import pandas as pd
import glob
import os

PASTA_DESCONTOS = r'C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\Descontos RGB'
PASTA_CUPONS    = r'C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\Cupons de Venda por SKU'
SAIDA           = r'C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\merge_bases_descontos\base_consolidada.csv'

KEY_ITEM = ['Codigo Loja', 'N. Boleto', 'Codigo SKU', 'Data Desconto']
KEY_ETAPA1 = KEY_ITEM + ['Origem Desconto', 'ID Campanha']

IGNORAR_TAGS = {
    'motivo desconto (subtotal)',
    '0 - não é desconto fidelidade',
    '0 - nao e desconto fidelidade',
    '0 - n\xe3o \xe9 desconto fidelidade',
    '',
}


def processar_arquivo(arq):
    """Etapa 1 + Etapa 2 vetorizadas para um arquivo CSV."""
    df = pd.read_csv(arq, sep=';', encoding='latin-1', decimal=',', dtype=str)

    # Converte numéricos
    for c in ['Valor Bruto', 'Valor Desconto', 'Valor Liquido']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].str.replace(',', '.'), errors='coerce').fillna(0.0)
    if 'Qtd.' in df.columns:
        df['Qtd.'] = pd.to_numeric(df['Qtd.'].str.replace(',', '.'), errors='coerce').fillna(0.0)

    # Garante colunas opcionais
    for col in ['Origem Desconto', 'ID Campanha', 'Motivo Desconto', 'Descricao Campanha']:
        if col not in df.columns:
            df[col] = ''
    df['Origem Desconto'] = df['Origem Desconto'].fillna('').str.strip().str.upper()

    # ── Etapa 1: pré-consolidação (elimina duplicidades dentro do arquivo) ──
    agg1 = {c: 'first' for c in df.columns if c not in KEY_ETAPA1}
    agg1.update({'Qtd.': 'sum', 'Valor Bruto': 'sum', 'Valor Desconto': 'sum', 'Valor Liquido': 'sum'})
    df = df.groupby(KEY_ETAPA1, sort=False, dropna=False).agg(agg1).reset_index()

    # ── Etapa 2: base por item — vetorizada ──────────────────────────────────

    # 2a. Campos estáticos: primeiro valor por grupo
    base = df.groupby(KEY_ITEM, sort=False).agg(
        Loja=('Loja', 'first'),
        Descricao_Produto=('Descricao Produto', 'first'),
        Qtd=('Qtd.', 'max'),
        Valor_Bruto=('Valor Bruto', 'max'),  # MAX = bruto real (não soma entre origens)
    ).reset_index()
    base.rename(columns={'Descricao_Produto': 'Descricao Produto', 'Valor_Bruto': 'Valor Bruto'}, inplace=True)

    # 2b. Descontos por origem: pivot_table vetorizado
    piv = df.pivot_table(
        index=KEY_ITEM,
        columns='Origem Desconto',
        values='Valor Desconto',
        aggfunc='sum',
        fill_value=0.0,
    ).reset_index()
    piv.columns.name = None
    for orig in ['FIDELIDADE', 'PROMOCIONAL', 'MANUAL']:
        if orig not in piv.columns:
            piv[orig] = 0.0
    piv = piv.rename(columns={
        'FIDELIDADE':  'Desc Fidelidade',
        'PROMOCIONAL': 'Desc Promocional',
        'MANUAL':      'Desc Manual',
    })
    keep_cols = KEY_ITEM + ['Desc Fidelidade', 'Desc Promocional', 'Desc Manual']
    piv = piv[[c for c in keep_cols if c in piv.columns]]

    # 2c. Campanhas: concatena strings únicas por grupo (vetorizado com transform)
    df_camp = df[df['Descricao Campanha'].notna() & (df['Descricao Campanha'].str.strip() != '')][
        KEY_ITEM + ['Descricao Campanha']
    ].drop_duplicates()
    if len(df_camp):
        camp = df_camp.groupby(KEY_ITEM, sort=False)['Descricao Campanha'].agg(
            lambda x: ' | '.join(sorted(x.unique()))
        ).reset_index(name='Campanhas')
    else:
        camp = base[KEY_ITEM].copy()
        camp['Campanhas'] = ''

    # 2d. Tags: filtra inválidas antes de agrupar
    df['_mot'] = df['Motivo Desconto'].fillna('')
    # Explode por ' / ' (separador do campo Motivo Desconto)
    tags_exp = df[KEY_ITEM + ['_mot']].copy()
    tags_exp['_mot'] = tags_exp['_mot'].str.split('/')
    tags_exp = tags_exp.explode('_mot')
    tags_exp['_mot'] = tags_exp['_mot'].str.strip()
    tags_exp = tags_exp[~tags_exp['_mot'].str.lower().isin(IGNORAR_TAGS)]
    tags_exp = tags_exp[tags_exp['_mot'] != ''].drop_duplicates()
    if len(tags_exp):
        tags = tags_exp.groupby(KEY_ITEM, sort=False)['_mot'].agg(
            lambda x: ' | '.join(sorted(x.unique()))
        ).reset_index(name='Tags Desconto')
    else:
        tags = base[KEY_ITEM].copy()
        tags['Tags Desconto'] = ''

    # Merge tudo
    result = base \
        .merge(piv,  on=KEY_ITEM, how='left') \
        .merge(camp, on=KEY_ITEM, how='left') \
        .merge(tags, on=KEY_ITEM, how='left')

    for c in ['Desc Fidelidade', 'Desc Promocional', 'Desc Manual']:
        result[c] = result[c].fillna(0.0)
    result[['Campanhas', 'Tags Desconto']] = result[['Campanhas', 'Tags Desconto']].fillna('')

    return result


# ─── 1. Processar todos os arquivos de Descontos RGB ─────────────────────────
print("=== ETAPA 1+2: Processando arquivos de Descontos RGB (vetorizado) ===")
arquivos = sorted([
    f for f in glob.glob(os.path.join(PASTA_DESCONTOS, '*.csv'))
    if not os.path.basename(f).startswith('descontos_consolidado')
    and not os.path.basename(f).startswith('concatenar')
])
print(f"Arquivos encontrados: {len(arquivos)}")

bases = []
for arq in arquivos:
    try:
        df2 = processar_arquivo(arq)
        bases.append(df2)
        print(f"  {os.path.basename(arq)}: {len(df2):,} itens")
    except Exception as e:
        print(f"  ERRO {os.path.basename(arq)}: {e}")

base_total = pd.concat(bases, ignore_index=True)
print(f"\nTotal antes dedup entre arquivos: {len(base_total):,}")
base_total = base_total.drop_duplicates(subset=KEY_ITEM)
print(f"Total após dedup: {len(base_total):,}")

# ─── 2. Chave para join ───────────────────────────────────────────────────────
base_total['_chave'] = (
    base_total['Codigo Loja'].astype(str).str.strip()
    + '-'
    + base_total['N. Boleto'].astype(str).str.strip()
    + '-'
    + base_total['Data Desconto'].astype(str).str.strip()
)

# ─── 3. Cupons: apenas Vendedor e Canal de venda ──────────────────────────────
print("\n=== ETAPA 3: Carregando Cupons (Vendedor + Canal de venda) ===")
cup_arquivos = sorted([
    f for f in glob.glob(os.path.join(PASTA_CUPONS, '*.csv'))
    if not os.path.basename(f).startswith('cupons_consolidado')
    and not os.path.basename(f).startswith('concatenar')
])
print(f"Arquivos de cupons: {len(cup_arquivos)}")

cup_dfs = []
for arq in cup_arquivos:
    try:
        df = pd.read_csv(arq, sep=';', encoding='latin-1', dtype=str)
        col_cancel = next((c for c in df.columns if 'Cancel' in c), None)
        col_loja   = next((c for c in df.columns if c.strip() == 'Loja'), None)
        col_data   = next((c for c in df.columns if c.strip() == 'Data'), None)
        col_doc    = next((c for c in df.columns if 'Doc' in c or 'doc' in c), None)
        col_canal  = next((c for c in df.columns if 'Canal' in c), None)
        col_vend   = next((c for c in df.columns if 'Vendedor' in c), None)
        if not all([col_loja, col_data, col_doc, col_canal, col_vend]):
            print(f"  AVISO {os.path.basename(arq)}: colunas faltando — ignorando")
            continue
        if col_cancel:
            df = df[df[col_cancel].astype(str).str.strip() == 'N']
        sub = df[[col_loja, col_data, col_doc, col_canal, col_vend]].copy()
        sub.columns = ['Loja', 'Data', 'NDoc', 'Canal de venda', 'Vendedor']
        sub['Cod Loja'] = sub['Loja'].str.strip().str.split(' - ').str[0]
        sub['_chave'] = sub['Cod Loja'] + '-' + sub['NDoc'].str.strip() + '-' + sub['Data'].str.strip()
        cup_dfs.append(sub[['_chave', 'Canal de venda', 'Vendedor']])
    except Exception as e:
        print(f"  ERRO {os.path.basename(arq)}: {e}")

cup = pd.concat(cup_dfs, ignore_index=True)
print(f"Linhas cupons: {len(cup):,}")
cup = cup.drop_duplicates(subset='_chave')
print(f"Cupons únicos: {len(cup):,}")

# ─── 4. Join ──────────────────────────────────────────────────────────────────
print("\n=== ETAPA 4: Join com Cupons ===")
result = base_total.merge(cup, on='_chave', how='left')
sem_match = result['Canal de venda'].isna().sum()
print(f"Linhas após join: {len(result):,}")
print(f"Sem match em Cupons: {sem_match:,} ({sem_match/len(result)*100:.1f}%)")

result['Nome Consultor'] = result['Vendedor'].str.split(' - ').str[1].str.strip()
result = result.drop(columns=['_chave', 'Vendedor'], errors='ignore')

# ─── 5. Salvar ────────────────────────────────────────────────────────────────
result.to_csv(SAIDA, sep=';', index=False, encoding='utf-8-sig')
print(f"\nSalvo: {SAIDA}")
print(f"  {len(result):,} linhas | {len(result.columns)} colunas: {list(result.columns)}")

print("\n=== TOTAIS (conferência) ===")
print(f"  Bruto total:       R$ {result['Valor Bruto'].sum():>15,.2f}")
print(f"  Desc Promocional:  R$ {result['Desc Promocional'].sum():>15,.2f}")
print(f"  Desc Manual:       R$ {result['Desc Manual'].sum():>15,.2f}")
print(f"  Desc Fidelidade:   R$ {result['Desc Fidelidade'].sum():>15,.2f}")
itens_c_manual = (result['Desc Manual'] > 0).sum()
print(f"  Itens c/ manual:   {itens_c_manual:,}")
print(f"  % itens c/ manual: {itens_c_manual/len(result)*100:.1f}%")
