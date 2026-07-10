# CLAUDE.md

Este arquivo fornece orientações ao Claude Code (claude.ai/code) ao trabalhar com o código neste repositório.

## Visão Geral

Este é um repositório de projetos de uso geral — uma coleção de projetos independentes, a maioria arquivos estáticos soltos na raiz, alguns em subpastas próprias quando têm mais de um arquivo. Não há sistema de build, runtime ou suite de testes compartilhado entre projetos; cada um é autocontido.

## Projetos

### `reconciliacao-fonte-primaria/`

Motor de reconciliação de folha que recalcula Folha e Contábil direto das
fontes primárias (FOPAG, BALANCETE), usando a aba CHECK como segundo ponto
de conferência, não como fonte. Projeto novo, independente de qualquer
painel de conciliação anterior — ver `reconciliacao-fonte-primaria/README.md`
para o princípio completo.

- **Build**: nenhum. Sem dependências de runtime; `vendor/exceljs.min.js` (ExcelJS) já está no repo.
- **Rodar**: abrir `reconciliacao-fonte-primaria/index.html` num navegador, ou servir a pasta com um servidor estático (`npx serve reconciliacao-fonte-primaria`).
- **Teste**: `cd reconciliacao-fonte-primaria && npm run test:all` (unidade sem dependência externa + integração contra fixture `.xlsx` fictícia, ver README do projeto).
- **Lint/formatação**: nenhuma ferramenta configurada.
- **Arquitetura**: `engine.js` é o motor puro (parser de fórmula, classificador de proveniência de célula, recalculadores, reconciliação 3 vias, score de confiabilidade, cobertura de contas) e não depende de biblioteca de planilha — recebe um `WorkbookAdapter` injetado. `app.js` é a camada de UI: lê o XLSX via ExcelJS, detecta abas/blocos de mês na CHECK, monta o adapter e chama o engine. `index.html` é upload + dashboard.

## Ao Adicionar Novos Projetos

Quando um projeto for adicionado ao repositório, documente-o nesta seção com:

1. **Comandos de build** — como instalar dependências e compilar/empacotar.
2. **Comandos de teste** — como executar a suite completa e um teste isolado.
3. **Comandos de lint/formatação** — as ferramentas utilizadas e como invocá-las.
4. **Arquitetura** — estrutura de alto nível, pontos de entrada e fluxos de dados que abrangem múltiplos arquivos.
