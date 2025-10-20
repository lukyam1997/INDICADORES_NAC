const chartInstances = new Map();

const chartPalette = {
  ambulatorial: {
    primary: "#2563eb",
    secondary: "#0f766e",
  },
  cirurgico: {
    primary: "#16a34a",
    secondary: "#f97316",
  },
};

const ensureChart = (id, ctx, dataset, data) => {
  const palette = chartPalette[dataset];
  if (!palette) {
    throw new Error(`Paleta inexistente para o dataset ${dataset}`);
  }

  const ChartJs = globalThis.Chart;
  if (!ChartJs) {
    console.warn("Chart.js não foi carregado");
    return;
  }

  const previous = chartInstances.get(id);
  if (previous) {
    previous.destroy();
  }

  const config = {
    type: "line",
    data: {
      labels: data.map((item) => item.mes),
      datasets: [
        {
          label: "Produção",
          data: data.map((item) => item.total),
          borderColor: palette.primary,
          backgroundColor: "transparent",
          tension: 0.38,
          borderWidth: 3,
          pointRadius: 5,
        },
        {
          label: "Meta",
          data: data.map((item) => item.meta),
          borderColor: palette.secondary,
          backgroundColor: "transparent",
          borderDash: [6, 4],
          tension: 0.24,
          borderWidth: 2,
          pointRadius: 0,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            usePointStyle: true,
            color: getComputedStyle(document.body).getPropertyValue(
              "--text-secondary"
            ),
          },
        },
        tooltip: {
          backgroundColor: getComputedStyle(document.body).getPropertyValue(
            "--bg-surface"
          ),
          titleColor: getComputedStyle(document.body).getPropertyValue(
            "--text-primary"
          ),
          bodyColor: getComputedStyle(document.body).getPropertyValue(
            "--text-secondary"
          ),
          borderWidth: 1,
          borderColor: getComputedStyle(document.body).getPropertyValue(
            "--border-default"
          ),
        },
      },
      scales: {
        y: {
          ticks: {
            color: getComputedStyle(document.body).getPropertyValue(
              "--text-tertiary"
            ),
          },
          grid: {
            color: getComputedStyle(document.body).getPropertyValue(
              "--border-default"
            ),
          },
        },
        x: {
          ticks: {
            color: getComputedStyle(document.body).getPropertyValue(
              "--text-tertiary"
            ),
          },
          grid: {
            display: false,
          },
        },
      },
    },
  };

  const chart = new ChartJs(ctx, config);
  chartInstances.set(id, chart);
};

export { ensureChart };
