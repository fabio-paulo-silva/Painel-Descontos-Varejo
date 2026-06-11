# Painel de Descontos — Varejo

Dashboard de auditoria de descontos (Manual, Promocional, Fidelidade) publicado via **GitHub Pages**.

🔗 **Link do painel:** https://fabio-paulo-silva.github.io/Painel-Descontos-Varejo/

## Como atualizar (diário/quando quiser)

1. Garanta que a base `merge_consolidado_geral.csv` e a `dCentros.xlsx` estão nesta pasta (atualizadas).
2. Dê **duplo clique** em `atualizar_painel.bat`.
   - Ele regenera o painel e envia para o GitHub; o Pages republica em ~1–2 min.
   - **O link nunca muda.**

## Estrutura

- `gen_dashboard_v2.py` — gera o painel a partir do CSV + dCentros.
- `dist/` — o que é publicado no Pages:
  - `index.html` — o dashboard (página inicial).
  - `painel_dados.js` — dados agregados (KPIs, gráficos).
  - `detalhe_base.js` — cadastro de lojas/produtos/consultores.
  - `detalhe_AAAA-MM.js` — cupons de cada mês (carregados sob demanda).

## Importante

- As bases brutas (`*.csv`, `*.xlsx`) **não** vão para o GitHub (são grandes e ficam só na sua máquina — ver `.gitignore`).
- Os dados publicados são **públicos** (qualquer pessoa com o link acessa).
- O detalhamento por cupom carrega **um mês por vez** — escala o ano inteiro sem travar.
