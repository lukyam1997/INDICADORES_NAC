import { createElement } from "../utils/dom.js";

const FILTER_OPTIONS = {
  ambulatorial: {
    especialidade: [
      "Todas",
      "Cardiologia",
      "Clínica Médica",
      "Endocrinologia",
      "Pediatria",
    ],
    profissional: ["Todos", "Dr. João Alves", "Dra. Fernanda Rocha", "Dra. Luiza Costa"],
  },
  cirurgico: {
    especialidade: ["Todas", "Geral", "Ortopedia", "Neurologia"],
    equipe: ["Todas", "Equipe 1", "Equipe 2", "Equipe 3"],
  },
};

const formatLabel = (value) =>
  value
    .replace(/([A-Z])/g, " $1")
    .replace(/_/g, " ")
    .replace(/^./, (char) => char.toUpperCase())
    .trim();

const buildSelect = (name, label, values, currentValue, onChange) => {
  const wrapper = createElement("div", { className: "filter-group" });
  const labelEl = createElement("span", {
    className: "filter-group__label",
    text: label,
    attrs: { id: `${name}-label` },
  });
  const selectWrapper = createElement("div", { className: "select" });
  const select = createElement("select", {
    attrs: {
      name,
      "aria-labelledby": `${name}-label`,
    },
  });

  values.forEach((value) => {
    const option = createElement("option", {
      text: value,
      attrs: { value },
    });
    if (value === currentValue) {
      option.selected = true;
    }
    select.append(option);
  });

  select.addEventListener("change", (event) => {
    onChange(name, event.target.value);
  });

  const icon = createElement("span", {
    className: "select__icon",
    html: "&#9662;",
    attrs: { "aria-hidden": "true" },
  });

  selectWrapper.append(select, icon);
  wrapper.append(labelEl, selectWrapper);
  return wrapper;
};

const createFilterPanel = (dataset, filters, onFilterChange) => {
  const config = FILTER_OPTIONS[dataset] ?? {};
  const container = createElement("div", { className: "panel" });
  container.append(
    createElement("div", {
      className: "panel__title",
      text: "Filtros avançados",
    }),
    createElement("p", {
      className: "panel__subtitle",
      text: "Refine a análise aplicando filtros combinados.",
    })
  );

  Object.entries(config).forEach(([key, values]) => {
    const current = filters[key] ?? values[0];
    container.append(
      buildSelect(key, formatLabel(key), values, current, onFilterChange)
    );
  });

  return container;
};

export { createFilterPanel };
