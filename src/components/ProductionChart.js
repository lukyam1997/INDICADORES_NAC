import { createElement } from "../utils/dom.js";
import { ensureChart } from "../utils/chartFactory.js";

const createProductionChart = (dataset, series) => {
  const container = createElement("div", { className: "panel" });
  container.append(
    createElement("div", {
      className: "panel__title",
      text: "Evolução mensal",
    }),
    createElement("p", {
      className: "panel__subtitle",
      text: "Produção real versus meta estabelecida pela gestão.",
    })
  );

  const chartWrapper = createElement("div", {
    className: "chart-grid__canvas",
    attrs: { style: "height: 320px" },
  });
  const canvas = createElement("canvas", { attrs: { id: "production-chart" } });
  chartWrapper.append(canvas);
  container.append(chartWrapper);

  requestAnimationFrame(() => {
    const ctx = canvas.getContext("2d");
    ensureChart("production-chart", ctx, dataset, series);
  });

  return container;
};

export { createProductionChart };
