const CONFIG = Object.freeze({
  spreadsheetId: "<SET YOUR SHEET ID HERE>",
  cacheTtlSeconds: 300,
  datasets: {
    ambulatorial: {
      sheetName: "AMBULATORIAL",
      headerRow: 1,
      range: "A1:Z",
      schema: {
        paciente: "Paciente",
        especialidade: "Especialidade",
        profissional: "Profissional",
        data: "Data",
        status: "Status",
        tempoAtendimento: "Tempo Atendimento (min)",
        resolutividade: "Resolutividade",
        retorno: "Retorno",
      },
    },
    cirurgico: {
      sheetName: "CIRURGICO",
      headerRow: 1,
      range: "A1:Z",
      schema: {
        paciente: "Paciente",
        procedimento: "Procedimento",
        equipe: "Equipe",
        data: "Data",
        status: "Status",
        diasInternacao: "Dias Internação",
        tempoEspera: "Tempo Espera",
        ocupacao: "Taxa Ocupação",
      },
    },
  },
});

class CacheStore {
  constructor(prefix) {
    this.prefix = prefix;
    this.cache = CacheService.getScriptCache();
  }

  buildKey(key) {
    return `${this.prefix}:${key}`;
  }

  get(key) {
    const cached = this.cache.get(this.buildKey(key));
    return cached ? JSON.parse(cached) : null;
  }

  set(key, value, ttlSeconds = CONFIG.cacheTtlSeconds) {
    this.cache.put(this.buildKey(key), JSON.stringify(value), ttlSeconds);
  }

  invalidate(prefix) {
    const keys = this.cache.getAllKeys();
    const scoped = keys.filter((key) => key.indexOf(this.prefix) === 0 && key.indexOf(prefix) !== -1);
    if (!scoped.length) return;
    this.cache.removeAll(scoped);
  }
}

class SheetGateway {
  constructor({ spreadsheetId, sheetName, range, headerRow }) {
    this.spreadsheetId = spreadsheetId;
    this.sheetName = sheetName;
    this.range = range;
    this.headerRow = headerRow;
  }

  fetchRawValues() {
    const spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(this.sheetName);
    if (!sheet) {
      throw new Error(`Planilha ${this.sheetName} não encontrada`);
    }
    const range = sheet.getRange(this.range);
    const values = range.getDisplayValues();
    const headers = values.shift();
    return { headers, rows: values.filter((row) => row.some((cell) => cell !== "")) };
  }
}

class DatasetMapper {
  constructor(schema) {
    this.schema = schema;
  }

  toRecord(headers, row) {
    const record = {};
    Object.entries(this.schema).forEach(([key, label]) => {
      const index = headers.indexOf(label);
      record[key] = index >= 0 ? row[index] : "";
    });
    return record;
  }
}

class DashboardRepository {
  constructor(gateway, mapper) {
    this.gateway = gateway;
    this.mapper = mapper;
  }

  list(filters) {
    const { headers, rows } = this.gateway.fetchRawValues();
    const mapped = rows.map((row) => this.mapper.toRecord(headers, row));
    return this.applyFilters(mapped, filters);
  }

  applyFilters(records, filters = {}) {
    const entries = Object.entries(filters).filter(([, value]) => value && value !== "Todas" && value !== "Todos");
    if (!entries.length) {
      return records;
    }
    return records.filter((record) =>
      entries.every(([key, value]) =>
        String(record[key]).toLowerCase().indexOf(String(value).toLowerCase()) !== -1
      )
    );
  }
}

class SummaryBuilder {
  static forDataset(dataset, records) {
    switch (dataset) {
      case "ambulatorial":
        return SummaryBuilder.buildAmbulatorial(records);
      case "cirurgico":
        return SummaryBuilder.buildCirurgico(records);
      default:
        throw new Error(`Dataset não suportado: ${dataset}`);
    }
  }

  static buildAmbulatorial(records) {
    const total = records.length;
    const tempo = SummaryBuilder.mean(records, "tempoAtendimento");
    const resolutividade = SummaryBuilder.mean(records, "resolutividade", 1);
    const retornos = records.filter((record) => `${record.retorno}`.toLowerCase() === "sim").length;

    return {
      atendimentos: total,
      tempoMedio: tempo,
      taxaResolutividade: resolutividade,
      retornos,
      tendencia: SummaryBuilder.trend(records, "tempoAtendimento"),
    };
  }

  static buildCirurgico(records) {
    const total = records.length;
    const ocupacao = SummaryBuilder.mean(records, "ocupacao", 1);
    const mediaPermanencia = SummaryBuilder.mean(records, "diasInternacao");
    const tempoEspera = SummaryBuilder.mean(records, "tempoEspera");

    return {
      procedimentos: total,
      ocupacao,
      mediaPermanencia,
      tempoEspera,
      tendencia: SummaryBuilder.trend(records, "tempoEspera", false),
    };
  }

  static mean(records, key, normalize = false) {
    const values = records
      .map((record) => SummaryBuilder.parseNumeric(record[key]))
      .filter((value) => !isNaN(value));
    if (!values.length) return 0;
    const sum = values.reduce((acc, value) => acc + value, 0);
    const result = sum / values.length;
    if (normalize) {
      const normalized = result > 1 ? result / 100 : result;
      return Number(normalized.toFixed(4));
    }
    return Number(result.toFixed(2));
  }

  static trend(records, key, inverse = true) {
    if (records.length < 2) return 0;
    const sorted = records
      .map((record) => ({
        data: new Date(record.data),
        value: SummaryBuilder.parseNumeric(record[key]),
      }))
      .filter((entry) => !isNaN(entry.value) && !isNaN(entry.data.getTime()))
      .sort((a, b) => a.data - b.data);

    if (sorted.length < 2) return 0;

    const last = sorted[sorted.length - 1].value;
    const penultimate = sorted[sorted.length - 2].value;
    if (penultimate === 0) return 0;

    const variation = (last - penultimate) / penultimate;
    return inverse ? Number((-variation).toFixed(2)) : Number(variation.toFixed(2));
  }

  static parseNumeric(value) {
    if (typeof value === "number") {
      return value;
    }
    if (typeof value === "string") {
      const sanitized = value
        .toString()
        .trim()
        .replace("%", "")
        .replace(/\./g, "")
        .replace(",", ".");
      return Number(sanitized);
    }
    return Number(value);
  }

  static monthLabel(month) {
    const months = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    return months[month - 1] || `${month}`;
  }
}

class SeriesBuilder {
  static forDataset(dataset, records) {
    const grouped = records.reduce((acc, record) => {
      const date = new Date(record.data);
      if (isNaN(date.getTime())) return acc;
      const key = `${date.getFullYear()}-${(`0${date.getMonth() + 1}`).slice(-2)}`;
      acc[key] = acc[key] || [];
      acc[key].push(record);
      return acc;
    }, {});

    return Object.keys(grouped)
      .sort()
      .map((key) => {
        const [year, month] = key.split("-");
        const total = grouped[key].length;
        return {
          mes: `${SummaryBuilder.monthLabel(Number(month))}/${year.slice(-2)}`,
          total,
          meta: SeriesBuilder.metaFor(dataset, key, total),
        };
      });
  }

  static metaFor(dataset, key, total) {
    const base = Math.max(total, 1);
    if (dataset === "ambulatorial") {
      return Math.round(base * 0.92);
    }
    return Math.round(base * 0.85);
  }
}

class DashboardService {
  constructor(config) {
    this.cache = new CacheStore("dashboard");
    this.config = config;
  }

  loadDataset(datasetKey, filters) {
    const dataset = this.config.datasets[datasetKey];
    if (!dataset) {
      throw new Error(`Dataset inválido: ${datasetKey}`);
    }

    const cacheKey = `${datasetKey}:${JSON.stringify(filters || {})}`;
    const cached = this.cache.get(cacheKey);
    if (cached) {
      return cached;
    }

    const gateway = new SheetGateway({
      spreadsheetId: this.config.spreadsheetId,
      sheetName: dataset.sheetName,
      range: dataset.range,
      headerRow: dataset.headerRow,
    });
    const mapper = new DatasetMapper(dataset.schema);
    const repository = new DashboardRepository(gateway, mapper);
    const records = repository.list(filters);

    const summary = SummaryBuilder.forDataset(datasetKey, records);
    const series = SeriesBuilder.forDataset(datasetKey, records);

    const response = { summary, series, records };
    this.cache.set(cacheKey, response);
    return response;
  }
}

class DashboardController {
  constructor(service) {
    this.service = service;
  }

  handle(request) {
    try {
      const dataset = request.dataset || "ambulatorial";
      const filters = request.filters || {};
      const snapshot = this.service.loadDataset(dataset, filters);
      const paginated = this.paginate(snapshot.records, request.page, request.pageSize);
      const body = {
        summary: snapshot.summary,
        series: snapshot.series,
        records: paginated.records,
        page: paginated.page,
        totalPages: paginated.totalPages,
        totalRecords: snapshot.records.length,
      };
      return this.json(body, 200);
    } catch (error) {
      console.error("Erro ao processar request", error);
      return this.json({ error: error.message }, 500);
    }
  }

  paginate(records, page = 1, pageSize = 25) {
    const currentPage = Math.max(Number(page) || 1, 1);
    const size = Math.max(Number(pageSize) || 25, 1);
    const offset = (currentPage - 1) * size;
    const paginated = records.slice(offset, offset + size);
    const totalPages = Math.max(Math.ceil(records.length / size), 1);
    return { records: paginated, page: currentPage, totalPages };
  }

  json(payload, statusCode) {
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON)
      .setStatusCode(statusCode);
  }
}

const controller = new DashboardController(new DashboardService(CONFIG));

function doPost(e) {
  const payload = parseRequest(e);
  return controller.handle(payload);
}

function parseRequest(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    console.error("Payload inválido", error);
    return {};
  }
}
