# Histórico de versões

Todas as mudanças relevantes do projeto. O formato segue a ideia do [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/).

## [1.1.0] · 19/10/2026

### Publicação

- Painel renomeado para `index.html` e publicado no GitHub Pages, com link direto no topo do README.
- Galeria com as oito telas da sala de situação e duas abas da análise completa, cada imagem ligada à tela correspondente no painel.
- Imagem de pré-visualização social (`docs/img/social-preview.png`).
- Guia de configuração do repositório (`docs/configuracao-github.md`).

### Base de dados

- Scripts para refazer a base a partir do painel: `scripts/extrair_dados.js` e `scripts/montar_base.py`.

## [1.0.0] · 19/10/2026

Primeira versão publicada, que reúne as etapas de desenvolvimento abaixo.

### Painel (`index.html`)

- **Etapa 1.** Painel com dados fictícios de um hospital de 230 leitos em dez unidades, com abas de censo hospitalar, indicadores e KPIs, giro de leitos, demanda do NIR e Kanban de permanência.
- **Etapa 2.** Painel de leitos em cards pequenos, com ficha do leito, filtros e busca.
- **Etapa 3.** Painel de desocupação: alta social, falta de transporte, falta de acompanhante, falta de leito de retaguarda e outros motivos, com custo detalhado por componente e receita de AIH não realizada.
- **Etapa 4.** Resumo comentado com tópicos clicáveis que abrem o painel detalhado de cada assunto.
- **Etapa 5.** Sala de situação com oito telas para TV, rodízio automático e navegação pelo teclado.
- **Etapa 6.** Estilo corporativo claro, no padrão de relatório Power BI.
- **Etapa 7.** Panorama geral com um cartão por unidade e balões de explicação ao clicar nos gráficos.

### Base de dados (`data/`)

- Base em modelo estrela com 13 tabelas, abas LEIA-ME, Conferencia e Dicionario com 115 colunas.
- Exportação das tabelas em CSV e extração bruta em JSON.

### Documentação

- README com insights da análise, plano de ação, medidas DAX, regras de negócio, referências e glossário.
