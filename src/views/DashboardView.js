import { createElement, clearElement } from "../utils/dom.js";
import { createFilterPanel } from "../components/FilterPanel.js";
import { createSummaryCards } from "../components/SummaryCards.js";
import { createProductionChart } from "../components/ProductionChart.js";
import { createRecordsTable } from "../components/RecordsTable.js";

const renderDashboard = (root, state, actions) => {
  clearElement(root);

  const { dataset, filters, summary, series, records, loading } = state;

  if (loading) {
    root.append(
      createElement("div", {
        className: "panel",
        children: [
          createElement("p", {
            className: "panel__title",
            text: "Carregando dados...",
          }),
        ],
      })
    );
    return;
  }

  const layout = createElement("div", { className: "dashboard-grid" });

  layout.append(
    createFilterPanel(dataset, filters, (name, value) => {
      actions.applyFilter(name, value);
    })
  );

  const mainColumn = createElement("div", { className: "dashboard-grid__content" });
  const summaryPanel = createSummaryCards(dataset, summary);
  const chartPanel = createProductionChart(dataset, series);
  const chartGrid = createElement("div", { className: "chart-grid" });
  chartGrid.append(chartPanel);

  mainColumn.append(summaryPanel, chartGrid, createRecordsTable(dataset, records));

  layout.append(mainColumn);
  root.append(layout);
};

export { renderDashboard };
