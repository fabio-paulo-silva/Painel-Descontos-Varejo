# Painel de Auditoria de Descontos — Documentação Completa

> **Para usar em outro chat:** cole este arquivo inteiro no início da conversa e diga o que deseja alterar.
> Link do painel publicado: https://fabio-paulo-silva.github.io/Painel-Descontos-Varejo/

---

## 1. Visão geral

Dashboard web de auditoria de descontos (Manual, Promocional, Fidelidade) para a Diretoria Comercial. Gerado por um script Python a partir de uma base CSV + cadastro de lojas (Excel), publicado automaticamente no GitHub Pages.

**Fluxo completo:**
```
merge_consolidado_geral.csv  ─┐
dCentros.xlsx                 ├─► gen_dashboard_v2.py ─► dist/ ─► GitHub Pages
```

---

## 2. Arquivos do projeto

**Pasta raiz:**
```
C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\merge_bases_descontos\
```

| Arquivo | Papel |
|---|---|
| `gen_dashboard_v2.py` | Script principal — lê CSV, agrega, gera todos os arquivos em `dist/` |
| `merge_consolidado_geral.csv` | Base bruta de vendas (≈2,1M linhas). **Não vai pro GitHub** |
| `dCentros.xlsx` | Cadastro de lojas: BCPS → Nome/Regional/Praça/Cluster/Gestores. **Não vai pro GitHub** |
| `atualizar_painel.bat` | Duplo clique → gera + commita + publica em 1 passo |
| `README.md` | Instruções para o usuário final |
| `PROJETO.md` | Este arquivo de documentação |
| `.gitignore` | Exclui `*.csv`, `*.xlsx` e arquivos temporários |
| `.github/workflows/pages.yml` | CI: deploy automático da pasta `dist/` no GitHub Pages |

**Pasta `dist/` (publicada no GitHub Pages):**

| Arquivo | Tamanho aprox. | Conteúdo |
|---|---|---|
| `index.html` | ~58 KB | Dashboard completo (HTML + CSS + JS inline) |
| `painel_dados.js` | ~20 MB | `window.PAINEL` — dados agregados: ROWS, ITEMS, lookups |
| `detalhe_base.js` | ~0,2 MB | `window.DET_BASE` — cadastros: lojas, produtos, consultores |
| `detalhe_2026-01.js` | ~12–17 MB | Cupons de Jan/2026 (carregado sob demanda) |
| `detalhe_2026-MM.js` | ~12–17 MB | Um arquivo por mês — gerado automaticamente |

---

## 3. Como atualizar o painel

```
Duplo clique em: atualizar_painel.bat
```

O bat faz:
1. `python gen_dashboard_v2.py` — lê o CSV e gera `dist/`
2. `git add dist ...` + `git commit`
3. `git push` → GitHub Pages republica em ~1–2 min

**Repositório GitHub:** `https://github.com/fabio-paulo-silva/Painel-Descontos-Varejo.git`
**Branch:** `main`
**Deploy:** `.github/workflows/pages.yml` (actions/upload-pages-artifact com `path: dist`)

---

## 4. Regras de negócio — cálculos críticos

### 4.1 Granularidade da base

Cada linha do CSV = 1 SKU × 1 origem de desconto dentro de um cupom.
O mesmo SKU pode ter múltiplas linhas (MANUAL + PROMOCIONAL + FIDELIDADE = 3 linhas por SKU).

**Regra do bruto:** `DESC_Valor Bruto` se **repete** em todas as linhas do mesmo SKU (empilhamento).
Logo: `bruto_sku = MAX(bruto das linhas)` — **nunca somar**.

**Chave de agregação:** `(Chave_Unica, SKU)` → 1 registro por SKU por cupom, com:
- `bruto = MAX(DESC_Valor Bruto)` das linhas
- `promo = SUM(DESC_Valor Desconto)` das linhas PROMOCIONAIS
- `manual = SUM(DESC_Valor Desconto)` das linhas MANUAIS
- `fidelidade = SUM(DESC_Valor Desconto)` das linhas FIDELIDADE

### 4.2 Fórmulas de % de desconto

| Métrica | Fórmula | Uso |
|---|---|---|
| % Manual | `manual / (bruto − promo)` | Denominador = Base_Manual (Seção 1.5 do plano de auditoria) |
| % Promo | `promo / bruto` | Sobre o bruto cheio |
| % Fidelidade | `fidelidade / bruto` | Sobre o bruto cheio |
| % Total | `(promo + manual + fidelidade) / bruto` | Sobre o bruto cheio |

### 4.3 Regras de auditoria

| Regra | Condição | Ação no painel |
|---|---|---|
| A | Canal 11 (Parcerias/Convênio) com desconto MANUAL | Coletada em VIOLATIONS |
| C | Manual sem justificativa (Motivo 2 vazio ou genérico) | Coletada em VIOLATIONS |
| F5 | Líquido real do item ≤ 0 (promo+manual+fid >= bruto) | `f5_cnt` em ROWS; tabela na aba Fidelidade |
| F60 | Desconto total > 60% do bruto | `f60_cnt` em ROWS; tabela EXCEPTIONS_F |

### 4.4 Exclusão de gestores

Gestores são identificados de dois modos:
1. Colunas GVO, GCVO/GPVO, GRVO da planilha `dCentros.xlsx`
2. Lista manual: `EXTRA_GESTOR_SHORT = {'THAIS SOARES', 'ANDREA MELONIO'}`
   - Comparação por nome CURTO: `primeiro + último sobrenome` do nome completo no CSV

Gestores são excluídos **somente** dos gráficos de análise por consultor (não das demais visões).

---

## 5. Estrutura de dados em memória

### 5.1 `window.PAINEL` (painel_dados.js)

```javascript
PAINEL = {
  ROWS_HEADER: ["mes","data","loja","consultor","canal","motivo2","campanha",
                "bruto","promo","manual","fidelidade","cnt","viol_a","viol_c",
                "f5_cnt","f60_cnt","dup_cnt"],
  ROWS: [[mes, data_iso, lojaIdx, consIdx, canalIdx, motivoIdx, campIdx,
          bruto, promo, manual, fidelidade, cnt, viol_a, viol_c, f5_cnt, f60_cnt, dup_cnt], ...],
  // Cada ROWS é agregado por (data, codloja, consultor, canal, motivo2, campanha)
  // lojaIdx/consIdx/etc = índices nos arrays LK.loja / LK.cons / etc.

  VIOLATIONS: [{loja, consultor, canal, data, regional, praca, cluster,
                manual, bruto, motivo2, regra}, ...],  // top 200 por valor manual desc
  EXCEPTIONS_F: [{loja, consultor, canal, data, regional, praca, cluster,
                  bruto, desc_total, pct, origem}, ...],  // top 200 por desc_total desc

  ITEMS: [[prodIdx, mes, bruto, promo, manual], ...],
  // Agregado por (mes, sku) — visão de item da REDE TODA
  // ⚠ LIMITAÇÃO ATUAL: sem dimensão loja/cluster → gráfico de itens só filtra por Mês
  ITEMS_PROD: ["NOME DO PRODUTO", ...],  // índice = prodIdx

  MES_LABELS: {"2026-01": "Jan/2026", ...},

  LK: {
    loja: ["Nome Loja 1", ...],         // índice = lojaIdx em ROWS
    cons: ["NOME CONS", ...],           // índice = consIdx em ROWS
    canal: ["Canal A", ...],
    motivo2: ["TAG X", ...],
    campanha: ["Campanha Y", ...],
    loja_reg: ["Regional A", ...],      // paralelo ao array loja
    loja_praca: ["Praça X", ...],       // paralelo ao array loja
    loja_cluster: ["Cluster 1", ...],   // paralelo ao array loja
    regionais: ["Regional A", ...],     // valores únicos ordenados
    pracas: ["Praça X", ...],
    clusters: ["Cluster 1", ...],
    cons_gestor: [0, 1, 0, ...],        // 1 = é gestor (excluir de gráficos de consultor)
  }
}
```

### 5.2 `window.DET_BASE` (detalhe_base.js)

```javascript
DET_BASE = {
  lojas: ["Nome Loja", ...],     // índice = lojaIdx em cupons
  cons: ["NOME CONS", ...],
  prod: ["NOME PRODUTO", ...],
  datas: ["DD/MM/YYYY", ...],
  loja_reg: [...], loja_praca: [...], loja_cluster: [...],
  regionais: [...], pracas: [...], clusters: [...],
  mot: ["TAG X", ...],           // motivoIdx em itens
  camp: ["Campanha Y", ...],     // campIdx em itens
  meses: ["2026-01", ...],
  mes_labels: {"2026-01": "Jan/2026", ...}
}
```

### 5.3 `window.DET_M["2026-MM"]` (detalhe_2026-MM.js, por mês)

```javascript
// Cada cupom:
[boleto, lojaIdx, dataIdx, consIdx, mes,
 bruto, promo, manual, fidelidade,
 items]  // array de itens abaixo

// Cada item dentro de items:
[sku, prodIdx, qtd, bruto, promo, manual, fidelidade, scans, motivoIdx, campIdx]
// scans > 1 = escaneamento duplicado (item digitado manual x vezes)
```

---

## 6. Funções JavaScript principais (index.html)

| Função | O que faz |
|---|---|
| `filteredRows()` | Filtra `ROWS` por todos os `activeFilters` (mes, data, loja, consultor, canal, motivo2, campanha, regional, praca, cluster) |
| `semGestores(rows)` | Remove linhas de consultores marcados como gestores (`LK.cons_gestor[idx] === 1`) |
| `topManualByPct(rows, groupIdx, topN)` | Ranking por `manual / (bruto − promo)`, com piso de faturamento |
| `topByPct(rows, groupIdx, descKey, topN)` | Ranking por `desc / bruto` para promo e fidelidade |
| `topItensByPct(metric, topN)` | Ranking de itens por %. **Só filtra por Mês** (ver limitação em §9) |
| `aggregate(rows, groupIdx, valueIdxs)` | Agrupa ROWS por índice e soma colunas |
| `avgLinePlugin(value, axis)` | Plugin Chart.js: linha tracejada **vermelha** (#FF4455) com rótulo "Média X,X%" |
| `hbar(id, labels, data, color, isPct, extra, avg)` | Gráfico de barras horizontais com rolagem |
| `vbar(id, labels, data, color, yIsPct, avg)` | Gráfico de barras verticais |
| `detPct(v)` | Formata % com cor: verde ≤30%, laranja >30%<50%, vermelho ≥50% |
| `buscarDetalhe()` | Filtra e exibe cupons da aba Detalhamento |
| `loadMonth(mes, cb)` | Lazy-load do arquivo `detalhe_YYYY-MM.js` |
| `populateFilters()` | Reconstrói a barra de filtros com estado atual |
| `switchTab(page)` | Troca aba e re-renderiza |

### Filtros ativos (objeto `activeFilters`)

```javascript
activeFilters = {
  mes: '',         // ex: '2026-03'
  loja: '',        // índice em LK.loja (string)
  consultor: '',   // índice em LK.cons (string)
  canal: '',       // índice em LK.canal (string)
  motivo2: '',     // índice em LK.motivo2 (string)
  campanha: '',    // índice em LK.campanha (string)
  dini: '',        // 'YYYY-MM-DD'
  dfim: '',        // 'YYYY-MM-DD'
  regional: '',    // string exata (nome da regional)
  praca: '',       // string exata (nome da praça)
  cluster: '',     // string exata (nome do cluster)
}
```

---

## 7. Paleta de cores

| Variável CSS | Hex | Uso |
|---|---|---|
| `--bg` | `#000000` | Fundo da página |
| `--card` | `#1C2B3A` | Cards, painéis |
| `--border` | `#2a3f54` | Bordas |
| `--text` | `#FFFFFF` | Texto principal |
| `--sub` | `#8aa3b8` | Texto secundário, labels |
| `--green` | `#2ECC8A` | Cor principal, KPI padrão, barras Manual |
| `--blue` | `#3355FF` | Barras Lojas |
| `--cyan` | `#00AAFF` | Barras Canal, KPIs |
| `--mint` | `#4DE8AA` | Destaques, % em tabelas |
| `--red` | `#FF4455` | Alertas, linha de média nos gráficos |
| `--yellow` | `#FFB830` | KPIs de % |
| `--navy` | `#1C2B3A` | Header |

**Cores de % no detalhamento:**
- Verde `--green` → ≤ 30%
- Laranja `#FF9F1C` → > 30% e < 50%
- Vermelho `--red` → ≥ 50%

---

## 8. Configurações do gerador Python

```python
BASE_DIR = r'C:\Users\fabio.silva\OneDrive - Gentil Negócios\Área de Trabalho\Fábio\merge_bases_descontos'
DIST     = os.path.join(BASE_DIR, 'dist')
CSV_PATH = os.path.join(BASE_DIR, 'merge_consolidado_geral.csv')
OUT_PATH = os.path.join(DIST, 'index.html')
DCENTROS_PATH = os.path.join(BASE_DIR, 'dCentros.xlsx')

# Gestoras cujo nome no CSV não casa com dCentros (mapeamento manual)
EXTRA_GESTOR_SHORT = {'THAIS SOARES', 'ANDREA MELONIO'}
```

**Colunas do CSV usadas (por índice variável — mapeado por nome no header):**

| Nome da coluna | Conteúdo |
|---|---|
| `DESC_Codigo Loja` | Código BCPS da loja (chave para dCentros) |
| `DESC_Loja` | Nome da loja (fallback se sem dCentros) |
| `DESC_Origem Desconto` | PROMOCIONAL / MANUAL / FIDELIDADE / SEM DESCONTOS |
| `DESC_Descricao Campanha` | Nome da campanha promocional |
| `DESC_Data Desconto` | DD/MM/YYYY |
| `DESC_Codigo SKU` | Código do produto |
| `DESC_Descricao Produto` | Nome do produto |
| `DESC_Qtd.` | Quantidade |
| `DESC_Valor Bruto` | Valor bruto do item (repete por origem!) |
| `DESC_Valor Desconto` | Valor do desconto desta linha/origem |
| `Chave Unica` | ID do cupom (ex: LOJA-BOLETO) |
| `CUP_Hora` | Hora do cupom |
| `CUP_Canal de venda` | Canal (ex: "11 - Parcerias") |
| `Motivo Desconto 2` | Justificativa do desconto manual (TAG) |
| `Nome Consultor` | Nome completo do consultor |

**Colunas do dCentros.xlsx usadas:**

| Coluna | Conteúdo |
|---|---|
| `BCPS` | Código da loja (chave de join com o CSV) |
| `LOJA` | Nome da loja para exibição |
| `REGIONAL` | Regional |
| `PRAÇA` (ou começa com PRA) | Praça |
| `CLUSTER` | Cluster |
| `GVO` | Nome do Gestor de Vendas Operacional |
| `GCVO/GPVO` | Nome do Gerente Comercial |
| `GRVO` | Nome do Gerente Regional |

---

## 9. Limitações conhecidas e melhorias pendentes

### 9.1 ⚠ Gráficos de ITENS não filtram por Loja/Cluster (pendente)

**Problema:** O array `ITEMS` em `painel_dados.js` agrega por `(mes, sku)` sem dimensão de loja.
A função `topItensByPct()` só filtra por `activeFilters.mes`.

**Impacto:** Ao aplicar filtro de Regional, Praça, Cluster ou Loja, os gráficos de barras
"Itens por % Manual" e "Itens por % Promo" **não mudam** — mostram sempre a visão da rede inteira.

**Como corrigir (requer rodar `gen_dashboard_v2.py`):**

*Python* — mudar a chave de `item_acc` de `(mes, sku)` para `(mes, sku, codloja)` e incluir `lojaIdx` no array ITEMS:
```python
# Linha ~273 do gen_dashboard_v2.py — ANTES:
k = (rec['mes'], rec['sku'])
ia = item_acc[k]
ia[0] += bruto; ia[1] += promo; ia[2] += manual

# DEPOIS:
k = (rec['mes'], rec['sku'], rec.get('codloja',''))
ia = item_acc[k]
ia[0] += bruto; ia[1] += promo; ia[2] += manual

# Linha ~283 — ANTES:
ITEMS.append([items_prod_idx[nm], mes, round(ia[0],2), round(ia[1],2), round(ia[2],2)])

# DEPOIS (adicionar lojaIdx como posição 1):
lojaIdx = _ix_loja_items(codloja)  # precisa de helper semelhante a _ix_loja
ITEMS.append([items_prod_idx[nm], mes, lojaIdx, round(ia[0],2), round(ia[1],2), round(ia[2],2)])
# e adicionar LK.items_loja (array de nomes de lojas para itens) no json de saída
```

*JavaScript* — atualizar `topItensByPct()` para filtrar por `activeFilters.loja/regional/praca/cluster`:
```javascript
function topItensByPct(metric, topN) {
  const mes = activeFilters.mes;
  const lojaF = activeFilters.loja !== '' ? +activeFilters.loja : -1;
  // Se tiver filtro de regional/praça/cluster, montar set de lojaIdx válidos
  // (requer LK.items_loja paralelo ao array ITEMS)
  const agg = {};
  for (let i = 0; i < ITEMS.length; i++) {
    const it = ITEMS[i];
    if (mes && it[1] !== mes) continue;
    // filtro de loja por índice (posição 2 no novo formato)
    if (lojaF >= 0 && it[2] !== lojaF) continue;
    // filtro por regional/praça/cluster via LK.items_loja[it[2]]
    const k = it[0];
    if (!agg[k]) agg[k] = {b:0, p:0, m:0};
    agg[k].b += it[3]; agg[k].p += it[4]; agg[k].m += it[5];
  }
  // ... resto igual
}
```

### 9.2 Filtro de Item/Produto (não implementado)

Adicionar campo pesquisável "Produto" na barra de filtros que filtre os gráficos de itens
para mostrar apenas aquele produto. Implementação sugerida:
- Adicionar `activeFilters.item = ''` (índice em `ITEMS_PROD`)
- No `topItensByPct()`, filtrar `if (activeFilters.item !== '' && k !== +activeFilters.item) continue`
- Na `populateFilters()`, adicionar `searchable('item', 'Produto', ITEMS_PROD)`

---

## 10. Aba Detalhamento — comportamento especial

- Não usa a barra de filtros global — tem seus próprios filtros internos
- **Mês é obrigatório** (select pré-selecionado no mais recente)
- **Loja (ou Regional/Praça/Cluster) é obrigatório**
- Lazy-load: carrega `detalhe_base.js` uma vez e `detalhe_YYYY-MM.js` sob demanda por mês
- Colunas ordináveis: click no cabeçalho inverte a ordem (padrão: %Total desc)
- Click na linha do cupom expande os produtos (toggle)
- Cada item mostra TAG (se desconto manual) e Campanha (se desconto promo)
- Escaneamento duplicado (scans > 1) exibe badge `⟳ Nx` em vermelho

---

## 11. Histórico de decisões importantes

| Decisão | Motivação |
|---|---|
| Bruto = MAX (não soma) | Bruto repete em cada linha de desconto do mesmo SKU — somar inflava 165M→105M |
| % Manual = manual/(bruto−promo) | Regra de alçada: base do manual é o valor já depois do desconto promo |
| Particionamento por mês no Detalhamento | Arquivo único cresceu 72MB (limite GitHub = 100MB); cada mês gera 12-17MB |
| Dados externos (painel_dados.js) | HTML fica ≤60KB e cacheável; dados regenerados sem alterar o HTML estrutural |
| Gestores excluídos dos gráficos de consultor | Vendas no nome do gestor são operacionais/transferências, distorcem % médio |
| EXTRA_GESTOR_SHORT | Nomes no dCentros não casam exatamente com CSV — mapeamento por nome curto |
| Strings armazenadas como índices em ROWS | Reduz painel_dados.js de ~80MB para ~20MB |
| Linha de média em vermelho tracejado | Pedido explícito para destacar a média vs. os indivíduos |

---

## 12. Como fazer alterações comuns

### Mudar a paleta de cores
Editar as variáveis CSS no início do template HTML dentro de `gen_dashboard_v2.py` (~linha 564):
```python
HTML = f"""...
:root {{
  --green: #2ECC8A;   ← trocar aqui
  --blue:  #3355FF;
  ...
}}
```

### Adicionar novo KPI
Na função `renderManual()` (ou renderPromo/renderFidelidade), após o bloco de KPIs existente:
```javascript
kpiCard(kpiEl, 'Rótulo', valor_formatado, 'subtítulo', 'colorClass');
// colorClasses disponíveis: '' (verde), 'yellow', 'blue', 'cyan', 'red'
```

### Adicionar nova coluna de filtro
1. Adicionar `novaCol: ''` em `activeFilters`
2. Em `filteredRows()`, adicionar a condição de filtro
3. Em `populateFilters()`, adicionar `searchable('novaCol', 'Rótulo', LK.novaCol)` ou `selDim('novaCol', 'Rótulo', LK.novaCol)`
4. No Python, incluir o lookup correspondente em `lookups_json`

### Incluir novo mês na base
Basta atualizar `merge_consolidado_geral.csv` com os novos dados e rodar `atualizar_painel.bat`.
O script detecta os meses presentes automaticamente e gera `detalhe_YYYY-MM.js` para cada um.

### Mudar limiar de cor no Detalhamento
Função `detPct(v)` (~linha 1730 do script Python, dentro do template HTML):
```javascript
function detPct(v) {
  if (v <= 0) return '<span class="zero">—</span>';
  const cls = v > 0.50 ? 'pctred' : (v > 0.30 ? 'pctorange' : 'pctgreen');
  // ↑ trocar 0.50 e 0.30 pelos novos limiares (em decimal)
  return '<span class="' + cls + '">' + (v*100).toFixed(1).replace('.', ',') + '%</span>';
}
```

### Adicionar gestora manualmente
No início de `gen_dashboard_v2.py`, linha ~70:
```python
EXTRA_GESTOR_SHORT = {'THAIS SOARES', 'ANDREA MELONIO', 'NOME SOBRENOME'}
# formato: primeiro nome + último sobrenome, em maiúsculas, sem acentos
```

---

## 13. Estrutura de pastas resumida

```
merge_bases_descontos/
├── gen_dashboard_v2.py        ← EDITAR AQUI para mudar lógica/visual
├── merge_consolidado_geral.csv  (não está no GitHub)
├── dCentros.xlsx                (não está no GitHub)
├── atualizar_painel.bat       ← duplo clique para publicar
├── PROJETO.md                 ← este arquivo
├── README.md
├── .gitignore
├── .github/
│   └── workflows/
│       └── pages.yml          ← deploy automático
└── dist/                      ← publicado no GitHub Pages
    ├── index.html
    ├── painel_dados.js
    ├── detalhe_base.js
    ├── detalhe_2026-01.js
    ├── detalhe_2026-02.js
    └── ...
```
