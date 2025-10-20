import { createElement } from "../utils/dom.js";

const buildHeaders = (dataset) => {
  if (dataset === "cirurgico") {
    return ["Paciente", "Procedimento", "Equipe", "Data", "Status"];
  }
  return ["Paciente", "Especialidade", "Profissional", "Data", "Status"];
};

const createRow = (dataset, record) => {
  const tr = document.createElement("tr");
  if (dataset === "cirurgico") {
    tr.append(
      createElement("td", { text: record.paciente }),
      createElement("td", { text: record.procedimento }),
      createElement("td", { text: record.equipe }),
      createElement("td", { text: formatDate(record.data) }),
      createElement("td", { text: record.status })
    );
  } else {
    tr.append(
      createElement("td", { text: record.paciente }),
      createElement("td", { text: record.especialidade }),
      createElement("td", { text: record.profissional }),
      createElement("td", { text: formatDate(record.data) }),
      createElement("td", { text: record.status })
    );
  }
  return tr;
};

const formatDate = (value) => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

const createRecordsTable = (dataset, records) => {
  const container = createElement("div", { className: "panel" });
  container.append(
    createElement("div", {
      className: "panel__title",
      text: "Atendimentos recentes",
    }),
    createElement("p", {
      className: "panel__subtitle",
      text: "Amostra das últimas movimentações registradas na base.",
    })
  );

  if (!records?.length) {
    container.append(
      createElement("div", {
        className: "table__empty",
        text: "Nenhum registro encontrado para os filtros selecionados.",
      })
    );
    return container;
  }

  const table = createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  buildHeaders(dataset).forEach((header) => {
    headerRow.append(createElement("th", { text: header }));
  });
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  records.forEach((record) => {
    tbody.append(createRow(dataset, record));
  });
  table.append(tbody);

  container.append(table);
  return container;
};

export { createRecordsTable };
