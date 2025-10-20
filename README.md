# Indicadores Gerenciados – Nova Arquitetura

Esta versão refatora totalmente o projeto anterior, separando responsabilidades entre uma aplicação front-end modular em ES Modules e uma API do Google Apps Script organizada em camadas. A seguir estão instruções de execução, estrutura de pastas e pontos de extensão.

## Estrutura do projeto

```
.
├── index.html              # Shell leve que injeta os estilos e o bundle modular
├── styles/
│   ├── tokens.css          # Design tokens e temas dinâmicos
│   ├── base.css            # Reset e regras tipográficas globais
│   ├── layout.css          # Layout e grid responsivos
│   └── components.css      # Componentes reutilizáveis
├── src/
│   ├── app.js              # Bootstrap da aplicação e orquestração de estado
│   ├── main.js             # Ponto de entrada do front-end
│   ├── state/              # Store reativa minimalista
│   ├── services/           # Consumo da API + dados de fallback
│   ├── components/         # UI componentizada (filtros, cards, gráficos, tabelas)
│   ├── utils/              # Helpers de DOM e Chart.js
│   └── views/              # Views de alto nível
└── Code.gs                 # API do Apps Script em camadas (cache, repositório, serviço, controller)
```

## Executando localmente

1. Sirva os arquivos com qualquer servidor estático (por exemplo `python -m http.server`).
2. Defina a variável global `INDICADORES_API_URL` para apontar para a URL de implantação do Apps Script, caso queira consumir dados reais.
3. Abra `http://localhost:8000` (ou a porta escolhida) no navegador.

A aplicação utilizará dados mock em caso de falhas de rede ou indisponibilidade da API.

## Publicação do Apps Script

1. Substitua `CONFIG.spreadsheetId` pelo ID correto da planilha no `Code.gs`.
2. Ajuste os nomes das abas e colunas dentro do objeto `CONFIG.datasets` se necessário.
3. Publique o projeto como aplicativo da web e habilite o método `doPost`.
4. Atualize a URL no front-end usando `INDICADORES_API_URL`.

O controller `DashboardController` responde a requisições POST no endpoint `/dashboard` com payload JSON (`dataset`, `filters`, `page`, `pageSize`).

## Customizações futuras

- **Novas visualizações**: crie arquivos adicionais em `src/components` e componha-os em `src/views/DashboardView.js`.
- **Novos datasets**: basta adicionar um novo item em `CONFIG.datasets` e configurar os schemas; o serviço automaticamente passa a suportá-lo.
- **Theming**: adicione tokens em `styles/tokens.css` e utilize as classes existentes para alternar temas.

## Testes sugeridos

- Testes unitários da camada de serviços podem ser escritos com `clasp` + `gas-local`.
- Para o front-end, recomenda-se adicionar `Vitest`/`Playwright` em uma etapa posterior de CI/CD.

Esta reescrita cria uma base modular, reutilizável e pronta para evoluções contínuas.
