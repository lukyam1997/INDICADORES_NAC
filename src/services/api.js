import { mockRecords, mockSeries, mockSummary } from "./mockData.js";

const BASE_URL =
  globalThis?.INDICADORES_API_URL ?? "https://script.google.com/macros/s/dummy/exec";

const withTimeout = async (promise, timeout = 6000) => {
  const timeoutPromise = new Promise((_, reject) => {
    const id = setTimeout(() => {
      clearTimeout(id);
      reject(new Error("Tempo limite excedido ao consultar os dados"));
    }, timeout);
  });

  return Promise.race([promise, timeoutPromise]);
};

const request = async (path, { dataset, filters, page, pageSize } = {}) => {
  const body = JSON.stringify({ dataset, filters, page, pageSize });
  try {
    const response = await withTimeout(
      fetch(`${BASE_URL}${path}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body,
      })
    );

    if (!response.ok) {
      throw new Error(`Erro na API (${response.status})`);
    }

    return await response.json();
  } catch (error) {
    console.warn("Falha na API, utilizando dados mock", error);
    return null;
  }
};

const loadDashboardSnapshot = async (dataset, filters) => {
  const payload = await request("/dashboard", { dataset, filters });
  if (payload) {
    return payload;
  }

  return {
    summary: mockSummary[dataset],
    series: mockSeries[dataset],
    records: mockRecords[dataset],
  };
};

export { loadDashboardSnapshot };
