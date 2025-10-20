import { createElement } from "../utils/dom.js";

const formatPercent = (value) => {
  const numeric = Number.isFinite(value) ? value : 0;
  return `${(numeric * 100).toFixed(1)}%`;
};
const formatTrend = (value) => {
  if (value === 0) return { text: "Estável", modifier: "" };
  const direction = value > 0 ? "up" : "down";
  const icon = value > 0 ? "▲" : "▼";
  return {
    text: `${icon} ${(Math.abs(value) * 100).toFixed(1)}% vs mês anterior`,
    modifier: direction,
  };
};

const SUMMARY_MAP = {
  ambulatorial: (
    { atendimentos = 0, tempoMedio = 0, taxaResolutividade = 0, retornos = 0, tendencia = 0 } = {},
    container
  ) => {
    container.append(
      createMetricCard("Atendimentos", atendimentos.toLocaleString("pt-BR"), tendencia),
      createMetricCard("Tempo médio (min)", tempoMedio, 0),
      createMetricCard("Taxa de resolutividade", formatPercent(taxaResolutividade), 0),
      createMetricCard("Retornos", retornos.toLocaleString("pt-BR"), 0)
    );
  },
  cirurgico: (
    { procedimentos = 0, ocupacao = 0, mediaPermanencia = 0, tempoEspera = 0, tendencia = 0 } = {},
    container
  ) => {
    container.append(
      createMetricCard("Procedimentos", procedimentos.toLocaleString("pt-BR"), tendencia),
      createMetricCard("Taxa de ocupação", formatPercent(ocupacao), 0),
      createMetricCard("Média de permanência (dias)", mediaPermanencia, 0),
      createMetricCard("Tempo de espera (dias)", tempoEspera, 0)
    );
  },
};

const createMetricCard = (label, value, trend) => {
  const { text, modifier } = formatTrend(trend);
  return createElement("div", {
    className: "metric-card",
    children: [
      createElement("span", {
        className: "metric-card__label",
        text: label,
      }),
      createElement("strong", {
        className: "metric-card__value",
        text: value,
      }),
      createElement("span", {
        className: `metric-card__trend ${modifier ? `is-${modifier}` : ""}`.trim(),
        text,
      }),
    ],
  });
};

const createSummaryCards = (dataset, summary) => {
  const container = createElement("div", { className: "panel" });
  container.append(
    createElement("div", {
      className: "panel__title",
      text: "Resumo executivo",
    }),
    createElement("p", {
      className: "panel__subtitle",
      text: "Principais indicadores acompanhados diariamente.",
    })
  );

  const grid = createElement("div", { className: "summary-grid" });
  const renderer = SUMMARY_MAP[dataset];
  if (renderer) {
    renderer(summary, grid);
  }

  container.append(grid);
  return container;
};

export { createSummaryCards };
