# Configuração do repositório no GitHub

Passo a passo para deixar o repositório `gestaodeleitos` com o painel interativo acessível pelo GitHub Pages, as telas no README e a imagem de pré-visualização. Tudo é feito pela interface do site.

## 1. Enviar os arquivos

1. Abra o repositório em `https://github.com/larissahorrara/gestaodeleitos`.
2. Clique em **Add file › Upload files**.
3. Arraste o **conteúdo** da pasta `gestaodeleitos` (não a pasta em si): `index.html`, `README.md`, `CHANGELOG.md`, `package.json`, `requirements.txt`, `.gitignore` e as pastas `data`, `docs` e `scripts`.
4. Clique em **Commit changes**. Os arquivos com o mesmo nome são substituídos.

> O arquivo `.gitignore` começa com ponto e pode não aparecer no explorador de arquivos. No Windows, ative "Itens ocultos" na guia **Exibir**; no Mac, use `Cmd + Shift + .`.

## 2. Apagar os arquivos da versão anterior

A primeira versão enviada tinha duas cópias que não fazem parte da estrutura final. Apague cada uma abrindo o arquivo, clicando no menu **⋯** e em **Delete file**, e depois em **Commit changes**:

- `Gestao_de_Leitos_Hospital_Regional_Horrara.html`, substituído pelo `index.html`;
- `base-dados-leitos.xlsx` na raiz, porque a base fica em `data/base-dados-leitos.xlsx`.

## 3. Publicar o painel no GitHub Pages

É esta etapa que torna o painel interativo pelo navegador. Sem ela, o GitHub mostra apenas o código do HTML.

1. Vá em **Settings › Pages**.
2. Em **Build and deployment › Source**, escolha **Deploy from a branch**.
3. Em **Branch**, selecione `main` e a pasta `/ (root)`. Clique em **Save**.
4. Aguarde alguns minutos. O painel fica disponível em `https://larissahorrara.github.io/gestaodeleitos/`.

O botão **Abrir o painel interativo**, a capa e as telas do README passam a abrir o painel quando esta etapa termina.

## 4. Menu "About" (lateral direita da página do repositório)

Clique na engrenagem ao lado de **About** e preencha:

**Description**

```text
Painel interativo de gestão de leitos hospitalares no padrão Power BI: censo, mapa de leitos, Kanban, desocupação com custo, fila do NIR e indicadores, com base em modelo estrela e medidas DAX. Dados fictícios.
```

**Website**

Marque **Use your GitHub Pages website**, que preenche `https://larissahorrara.github.io/gestaodeleitos/`.

**Topics** (digite um por vez e tecle Enter)

```text
gestao-de-leitos
gestao-hospitalar
nir
censo-hospitalar
kanban
saude-publica
sus
power-bi
dax
dashboard
data-visualization
business-intelligence
d3js
html
portfolio
```

## 5. Imagem de pré-visualização social

É a imagem que aparece quando o link do repositório é compartilhado no LinkedIn, WhatsApp e outras redes.

1. Vá em **Settings › General › Social preview** e clique em **Edit › Upload an image**.
2. Envie `docs/img/social-preview.png` (1280 × 640 pixels).

## 6. Repositório template (opcional)

1. Vá em **Settings › General**.
2. Logo abaixo do nome do repositório, marque **Template repository**.

A partir daí, aparece o botão **Use this template**, que gera uma cópia limpa do projeto.

## 7. Conferência final

- [ ] O README aparece com a capa, os selos e a galeria de telas.
- [ ] Na raiz do repositório não há mais o HTML com o nome antigo nem a planilha solta.
- [ ] O link `https://larissahorrara.github.io/gestaodeleitos/` abre o painel e os gráficos respondem ao clique.
- [ ] Clicar em uma tela da galeria abre a página correspondente do painel.
- [ ] O menu **About** mostra descrição, link e tópicos.
