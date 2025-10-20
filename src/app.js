import { createElement } from "./utils/dom.js";
import { createStore } from "./state/store.js";
import { renderDashboard } from "./views/DashboardView.js";
import { loadDashboardSnapshot } from "./services/api.js";

const initialState = {
  dataset: "ambulatorial",
  theme: "ambulatorial",
  filters: {},
  summary: null,
  series: [],
  records: [],
  loading: true,
};

const bootstrapApp = (root) => {
  const store = createStore(initialState);

  const header = buildHeader();
  const content = createElement("main", { className: "app__content" });

  root.append(header.element, content);

  const actions = createActions(store);

  store.subscribe((state) => {
    syncBodyAttributes(state);
    updateHeader(header, state, actions);
    renderDashboard(content, state, actions);
  });

  actions.refreshData();
};

const syncBodyAttributes = ({ dataset, theme }) => {
  document.body.dataset.dataset = dataset;
  document.body.dataset.theme = theme;
};

const createActions = (store) => {
  const refreshData = async () => {
    store.setState({ loading: true });
    const { dataset, filters } = store.getState();
    const data = await loadDashboardSnapshot(dataset, filters);
    store.setState({
      summary: data.summary,
      series: data.series,
      records: data.records,
      loading: false,
    });
  };

  return {
    refreshData,
    changeDataset: (dataset) => {
      store.setState({ dataset, filters: {} });
      refreshData();
    },
    toggleTheme: (theme) => {
      store.setState({ theme });
    },
    applyFilter: (name, value) => {
      store.setState((prev) => ({
        filters: {
          ...prev.filters,
          [name]: value,
        },
      }));
      refreshData();
    },
  };
};

const buildHeader = () => {
  const element = createElement("header", { className: "app__header" });

  const brand = createElement("div", {
    className: "app__brand",
    children: [
      createElement("span", {
        className: "app__brand-title",
        text: "Hospital Universitário do Ceará",
      }),
      createElement("span", {
        className: "app__brand-subtitle",
        text: "Indicadores gerenciados em tempo real",
      }),
    ],
  });

  const datasetChips = createChipGroup([
    { id: "ambulatorial", label: "Ambulatorial" },
    { id: "cirurgico", label: "Cirúrgico" },
  ]);

  const themeChips = createChipGroup([
    { id: "ambulatorial", label: "Tema Claro" },
    { id: "plantao", label: "Tema Plantão" },
  ]);

  element.append(brand, datasetChips.group, themeChips.group);

  return {
    element,
    datasetChips,
    themeChips,
  };
};

const createChipGroup = (options) => {
  const group = createElement("div", { className: "app__chip-group" });
  const buttons = options.map((option) => {
    const button = createElement("button", {
      className: "app__chip-button",
      text: option.label,
      attrs: { type: "button", "data-value": option.id },
    });
    group.append(button);
    return button;
  });
  return { group, buttons };
};

const updateHeader = ({ datasetChips, themeChips }, state, actions) => {
  datasetChips.buttons.forEach((button) => {
    const value = button.getAttribute("data-value");
    button.classList.toggle("is-active", value === state.dataset);
    button.onclick = () => actions.changeDataset(value);
  });

  themeChips.buttons.forEach((button) => {
    const value = button.getAttribute("data-value");
    button.classList.toggle("is-active", value === state.theme);
    button.onclick = () => actions.toggleTheme(value);
  });
};

export { bootstrapApp };
