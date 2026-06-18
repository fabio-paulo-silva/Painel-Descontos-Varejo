import pandas as pd, glob, os

pasta = r'C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\Cupons de Venda por SKU'
arquivos = [f for f in glob.glob(os.path.join(pasta, '*.csv'))
            if not os.path.basename(f).startswith('cupons_consolidado')
            and not os.path.basename(f).startswith('concatenar')]

dfs = []
for arq in arquivos:
    try:
        df = pd.read_csv(arq, sep=';', encoding='latin1', dtype=str)
        col_cancel = [c for c in df.columns if 'Cancel' in c or 'cancel' in c]
        col_data = [c for c in df.columns if c.strip() == 'Data']
        if not col_data:
            print('SEM COLUNA DATA: ' + os.path.basename(arq) + ' | colunas: ' + str(list(df.columns[:5])))
            continue
        if col_cancel:
            df = df[df[col_cancel[0]].astype(str).str.strip() == 'N']
        df['_arquivo'] = os.path.basename(arq)
        df['_data'] = pd.to_datetime(df[col_data[0]].str.strip(), dayfirst=True, errors='coerce')
        dfs.append(df[['_arquivo', '_data']])
    except Exception as e:
        print('ERRO ' + os.path.basename(arq) + ': ' + str(e))

if not dfs:
    print('Nenhum arquivo lido')
else:
    tudo = pd.concat(dfs, ignore_index=True)
    print('=== DATAS POR ARQUIVO ===')
    resumo = tudo.groupby('_arquivo')['_data'].agg(['min', 'max', 'count']).reset_index()
    resumo.columns = ['arquivo', 'min_data', 'max_data', 'linhas']
    resumo = resumo.sort_values('min_data')
    for _, r in resumo.iterrows():
        mn = r['min_data'].strftime('%d/%m') if pd.notna(r['min_data']) else 'N/A'
        mx = r['max_data'].strftime('%d/%m') if pd.notna(r['max_data']) else 'N/A'
        print(r['arquivo'] + ': ' + mn + ' a ' + mx + ' (' + str(r['linhas']) + ' itens)')

    print()
    print('=== COBERTURA GERAL ===')
    datas = tudo['_data'].dropna().dt.date
    todos_dias = pd.date_range(datas.min(), datas.max(), freq='D').date
    presentes = set(datas.unique())
    faltando = [d for d in todos_dias if d not in presentes]

    print('De: ' + str(datas.min()) + ' a ' + str(datas.max()))
    print('Dias esperados: ' + str(len(todos_dias)) + ' | Dias presentes: ' + str(len(presentes)))
    if faltando:
        print('DIAS FALTANDO: ' + str([str(d) for d in faltando]))
    else:
        print('Nenhum dia faltando.')

    # Verificar sobreposição entre arquivos
    print()
    print('=== SOBREPOSICAO ENTRE ARQUIVOS ===')
    por_data_arq = tudo.groupby(['_data', '_arquivo']).size().reset_index(name='n')
    datas_em_multiplos = por_data_arq.groupby('_data')['_arquivo'].count()
    datas_duplas = datas_em_multiplos[datas_em_multiplos > 1]
    if len(datas_duplas) == 0:
        print('Nenhuma data aparece em mais de um arquivo.')
    else:
        print('Datas em multiplos arquivos:')
        for dt in datas_duplas.index:
            arqs = por_data_arq[por_data_arq['_data'] == dt]['_arquivo'].tolist()
            print('  ' + str(dt.date()) + ': ' + ', '.join(arqs))
