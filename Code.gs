// =============================================
// DASHBOARD MÉDICO - GOOGLE APPS SCRIPT
// Otimizado para 40,000+ registros
// =============================================

// CONFIGURAÇÃO PRINCIPAL
const DATASET_DEFINITIONS = Object.freeze({
  ambulatorial: {
    id: 'ambulatorial',
    sheetName: 'BASE-DE-DADOS',
    displayName: 'Ambulatorial',
    themeSheetName: 'tema',
    columnCount: 51
  },
  cirurgico: {
    id: 'cirurgico',
    sheetName: 'BASE-DE-DADOS-2',
    displayName: 'Cirúrgico',
    themeSheetName: 'tema2',
    columnCount: 34
  }
});

const DEFAULT_DATASET_ID = 'ambulatorial';

const CONFIG = {
  spreadsheetId: '1E9sgdHUS6T6VFVGGP_FFhcqEYEYvJ8uyrN-QPbgRKQ0',
  sheetNames: {
    database: DATASET_DEFINITIONS[DEFAULT_DATASET_ID].sheetName,
    tema: 'tema'
  },
  datasets: DATASET_DEFINITIONS,
  maxRecords: 50000,
  cacheDuration: 600, // 10 minutos em segundos
  pageSize: 100, // Registros por página
  chunkSize: 1000, // Processamento em lotes
  baseDataCacheDuration: 600,
  baseDataCacheChunkSize: 150,
  baseDataCacheMaxChunkBytes: 90000
};

const BASE_DATA_CACHE = {
  version: 'v1'
};

const CACHE_REGISTRY_KEY = 'cacheRegistry_v1';
const CACHE_REGISTRY_MAX_KEYS_PER_PREFIX = 30;

function normalizeDatasetId(datasetId) {
  if (!datasetId) {
    return DEFAULT_DATASET_ID;
  }

  const normalized = datasetId.toString().trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(CONFIG.datasets, normalized)
    ? normalized
    : DEFAULT_DATASET_ID;
}

function getDatasetDefinition(datasetId) {
  const normalized = normalizeDatasetId(datasetId);
  return CONFIG.datasets[normalized];
}

function getDatasetSheet(spreadsheet, datasetId) {
  const definition = getDatasetDefinition(datasetId);
  return spreadsheet.getSheetByName(definition.sheetName);
}

function getDatasetThemeSheetName(datasetId) {
  const definition = getDatasetDefinition(datasetId);
  if (definition && definition.themeSheetName) {
    return definition.themeSheetName;
  }

  if (CONFIG.sheetNames && CONFIG.sheetNames.tema) {
    return CONFIG.sheetNames.tema;
  }

  return null;
}

function getDatasetTemaSheet(spreadsheet, datasetId) {
  const sheetName = getDatasetThemeSheetName(datasetId);
  if (!sheetName) {
    return null;
  }

  return spreadsheet.getSheetByName(sheetName);
}

function resolveDatasetIdFromFilters(filters) {
  let resolved = DEFAULT_DATASET_ID;

  if (filters && typeof filters === 'object') {
    const candidate = filters.dataset || filters.dataSet || filters.bi;
    if (candidate) {
      resolved = candidate;
    }
  }

  const normalized = normalizeDatasetId(resolved);
  try {
    console.log('[Dataset] resolveDatasetIdFromFilters', {
      filters: filters || {},
      resolved: normalized
    });
  } catch (error) {
    Logger.log('[Dataset] resolveDatasetIdFromFilters erro ao registrar log: %s', error);
  }

  return normalized;
}

function getBaseDataMetaKey(datasetId) {
  const normalized = normalizeDatasetId(datasetId);
  return `baseData_${normalized}_meta_${BASE_DATA_CACHE.version}`;
}

function buildBaseDataChunkKey(datasetId, timestamp, index) {
  const normalized = normalizeDatasetId(datasetId);
  return `baseData_${normalized}_chunk_${timestamp}_${index}`;
}

function sortFilterObjectForCache(filters) {
  const sorted = {};
  const source = filters || {};

  Object.keys(source).sort().forEach(key => {
    const value = source[key];
    if (Array.isArray(value)) {
      sorted[key] = value.slice().sort();
    } else {
      sorted[key] = value;
    }
  });

  return sorted;
}

/**
 * INICIALIZAÇÃO DO WEBAPP
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setTitle('Dashboard Médico - Gestão de Atendimentos');
}

/**
 * INCLUIR ARQUIVOS HTML/CSS/JS
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function sanitizeStringValue(value) {
  if (value === undefined || value === null) {
    return '';
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  if (typeof value === 'number') {
    return value.toString();
  }

  if (value.toString) {
    return value.toString().trim();
  }

  return String(value).trim();
}

function normalizeBooleanFlag(value) {
  if (value === true || value === 1) {
    return true;
  }

  if (value === false || value === 0) {
    return false;
  }

  if (typeof value === 'string') {
    const normalized = value.trim().toUpperCase();
    if (['SIM', 'S', 'YES', 'TRUE'].includes(normalized)) {
      return true;
    }

    if (['NÃO', 'NAO', 'N', 'NO', 'FALSE'].includes(normalized)) {
      return false;
    }
  }

  return false;
}

function mapAmbulatorialRowToRecord(row) {
  const safeRow = Array.isArray(row) ? row : [];

  return {
    prontuario: sanitizeStringValue(safeRow[2]),
    paciente: sanitizeStringValue(safeRow[4]),
    nascimento: safeRow[5] || '',
    profissional: sanitizeStringValue(safeRow[6]),
    tipo: sanitizeStringValue(safeRow[8]),
    especialidade: sanitizeStringValue(safeRow[9]),
    generoColJ: sanitizeStringValue(safeRow[10]),
    situacao: sanitizeStringValue(safeRow[11]),
    generoColL: sanitizeStringValue(safeRow[12]),
    generoColM: sanitizeStringValue(safeRow[13]),
    ano: sanitizeStringValue(safeRow[14]),
    mes: sanitizeStringValue(safeRow[15]),
    dataReferencia: safeRow[18] || '',
    quimioterapia: normalizeBooleanFlag(safeRow[24]),
    cirurgia: normalizeBooleanFlag(safeRow[26]),
    turno: sanitizeStringValue(safeRow[29])
  };
}

const CIRURGICO_FIELD_MAP = Object.freeze([
  { index: 0, keys: ['ANO'], transform: sanitizeStringValue },
  { index: 1, keys: ['MÊS', 'MES'], transform: sanitizeStringValue },
  { index: 2, keys: ['DATA DA SOLICITAÇÃO', 'DATA_DA_SOLICITACAO'], transform: value => value || '' },
  { index: 3, keys: ['PRONTUÁRIO', 'PRONTUARIO', 'PRONT'], transform: sanitizeStringValue },
  { index: 4, keys: ['NOME', 'PACIENTE'], transform: sanitizeStringValue },
  { index: 5, keys: ['GÊNERO', 'GENERO'], transform: sanitizeStringValue },
  { index: 6, keys: ['DATA DE NASCIMENTO', 'DATA_DE_NASCIMENTO'], transform: value => value || '' },
  { index: 7, keys: ['IDADE'], transform: value => (value === undefined || value === null ? '' : value) },
  { index: 8, keys: ['MUNICÍPIO', 'MUNICIPIO'], transform: sanitizeStringValue },
  { index: 9, keys: ['ESPECIALIDADE'], transform: sanitizeStringValue },
  { index: 10, keys: ['PROCEDIMENTO INDICADO', 'PROCEDIMENTO', 'PROCEDIMENTO_INDICADO'], transform: sanitizeStringValue },
  { index: 11, keys: ['CIRURGIÃO', 'CIRURGIAO'], transform: sanitizeStringValue },
  {
    index: 12,
    keys: [
      'CONFIRMAÇÃO DA INTERNAÇÃO',
      'CONFIRMAÇÃO_DA_INTERNAÇÃO',
      'CONFIRMACAO_DA_INTERNAÇÃO',
      'CONFIRMACAO_DA_INTERNACAO'
    ],
    transform: sanitizeStringValue
  },
  { index: 13, keys: ['DATA DE INTERNAÇÃO', 'DATA_DA_INTERNAÇÃO', 'DATA_DA_INTERNACAO'], transform: value => value || '' },
  { index: 14, keys: ['DATA DA CIRURGIA', 'DATA_DA_CIRURGIA'], transform: value => value || '' },
  { index: 15, keys: ['PROGRAMA'], transform: sanitizeStringValue },
  { index: 16, keys: ['STATUS DO AGENDAMENTO', 'STATUS_DO_AGENDAMENTO'], transform: sanitizeStringValue },
  { index: 17, keys: ['PENDÊNCIA 01', 'PENDENCIA_01'], transform: sanitizeStringValue },
  { index: 18, keys: ['PENDÊNCIA 02', 'PENDENCIA_02'], transform: sanitizeStringValue },
  { index: 19, keys: ['PENDÊNCIA 03', 'PENDENCIA_03'], transform: sanitizeStringValue },
  { index: 20, keys: ['SOLICITAÇÃO VAGA NA UTI?', 'SOLICITACAO_VAGA_NA_UTI'], transform: sanitizeStringValue },
  {
    index: 21,
    keys: ['COMORBIDADE? MEDICAMENTOS DE ROTINA ?', 'COMORBIDADE_MEDICAMENTOS'],
    transform: sanitizeStringValue
  },
  { index: 22, keys: ['VAD (VIAS AÉREAS DIFÍCEIS)', 'VAD'], transform: sanitizeStringValue },
  { index: 23, keys: ['CONGELAÇÃO', 'CONGELACAO'], transform: sanitizeStringValue },
  {
    index: 24,
    keys: ['TEMPO DE ESPERA PARA AGENDAMENTO CIRURGICO', 'TEMPO_DE_ESPERA'],
    transform: value => (value === undefined || value === null ? '' : value)
  },
  {
    index: 25,
    keys: ['HISTOPATOLOGICO COMPROVANDO NEOPLASIA', 'HISTOPATOLOGICO'],
    transform: sanitizeStringValue
  },
  { index: 26, keys: ['N° FASTMEDIC', 'N_FASTMEDIC'], transform: sanitizeStringValue },
  {
    index: 27,
    keys: ['classificação em fila DA ORTOPEDIA', 'CLASSIFICACAO_FILA_ORTOPEDIA'],
    transform: sanitizeStringValue
  },
  {
    index: 28,
    keys: ['SITUAÇÃO \nFASTMEDIC', 'SITUAÇÃO \u000AFASTMEDIC', 'SITUAÇÃO FASTMEDIC', 'SITUAÇÃO_FASTMEDIC', 'SITUACAO_FASTMEDIC'],
    transform: sanitizeStringValue
  },
  {
    index: 29,
    keys: ['SITUAÇÃO DO FORMULARIO DE AUTORIZAÇÃO DO MES', 'SITUACAO_FORM_AUTORIZACAO'],
    transform: sanitizeStringValue
  },
  { index: 30, keys: ['OBSERVAÇÃO 01', 'OBSERVACAO_01'], transform: sanitizeStringValue },
  { index: 31, keys: ['OBSERVAÇÃO 02', 'OBSERVACAO_02'], transform: sanitizeStringValue },
  { index: 32, keys: ['OBSERVAÇÃO 03', 'OBSERVACAO_03'], transform: sanitizeStringValue },
  { index: 33, keys: ['OBSERVAÇÃO 04', 'OBSERVACAO_04'], transform: sanitizeStringValue }
]);

function mapCirurgicoRowToRecord(row) {
  const safeRow = Array.isArray(row) ? row : [];
  const record = {};

  CIRURGICO_FIELD_MAP.forEach(field => {
    const transform = typeof field.transform === 'function' ? field.transform : sanitizeStringValue;
    const rawValue = safeRow[field.index];
    const value = transform(rawValue);

    field.keys.forEach(key => {
      if (!key) {
        return;
      }
      record[key] = value;
    });
  });

  return record;
}

function mapRowToRecord(row) {
  return mapAmbulatorialRowToRecord(row);
}

function mapDatasetRowToRecord(datasetId, row) {
  const normalized = normalizeDatasetId(datasetId);
  if (normalized === 'cirurgico') {
    return mapCirurgicoRowToRecord(row);
  }
  return mapAmbulatorialRowToRecord(row);
}

function getDatasetColumnCount(datasetId, sheet) {
  const normalized = normalizeDatasetId(datasetId);
  const definition = CONFIG.datasets[normalized] || DATASET_DEFINITIONS[normalized];

  if (definition && Number.isInteger(definition.columnCount) && definition.columnCount > 0) {
    return definition.columnCount;
  }

  if (sheet && typeof sheet.getLastColumn === 'function') {
    return sheet.getLastColumn();
  }

  return 30;
}

function coerceRowToRecord(row) {
  if (!row) {
    return mapAmbulatorialRowToRecord([]);
  }

  if (Array.isArray(row)) {
    return mapAmbulatorialRowToRecord(row);
  }

  if (typeof row === 'object') {
    if (row.prontuario !== undefined || row.situacao !== undefined) {
      return row;
    }

    const prontuario = sanitizeStringValue(row.prontuario
      || row.PRONTUARIO
      || row['PRONTUÁRIO']
      || row.PRONT);
    const paciente = sanitizeStringValue(row.paciente || row.NOME);
    const profissional = sanitizeStringValue(row.profissional || row.CIRURGIAO || row['CIRURGIÃO']);
    const tipo = sanitizeStringValue(row.tipo || row['PROCEDIMENTO INDICADO'] || row.PROCEDIMENTO_INDICADO);
    const especialidade = sanitizeStringValue(row.especialidade || row.ESPECIALIDADE);
    const situacao = sanitizeStringValue(row.situacao
      || row.SITUACAO
      || row['SITUAÇÃO_FASTMEDIC']
      || row['SITUACAO_FASTMEDIC']);
    const ano = sanitizeStringValue(row.ano || row.ANO);
    const mes = sanitizeStringValue(row.mes || row.MES || row['MÊS']);
    const dataReferencia = row.dataReferencia || row['DATA DA CIRURGIA'] || row.DATA_DA_CIRURGIA || '';
    const quimioterapia = normalizeBooleanFlag(row.quimioterapia);
    const cirurgia = normalizeBooleanFlag(row.cirurgia || !!row['DATA DA CIRURGIA'] || !!row.DATA_DA_CIRURGIA);
    const turno = sanitizeStringValue(row.turno || '');

    if (prontuario || paciente || profissional || tipo || especialidade || situacao || ano || mes || dataReferencia) {
      return {
        prontuario,
        paciente,
        profissional,
        tipo,
        especialidade,
        situacao,
        ano,
        mes,
        dataReferencia,
        quimioterapia,
        cirurgia,
        turno
      };
    }
  }

  return mapAmbulatorialRowToRecord([]);
}

function sanitizeFilters(filters) {
  if (!filters || typeof filters !== 'object') {
    return {};
  }

  const ignoredKeys = new Set(['page', 'pageSize']);

  return Object.keys(filters).reduce((acc, key) => {
    if (ignoredKeys.has(key)) {
      return acc;
    }

    const value = filters[key];

    if (Array.isArray(value)) {
      const sanitizedArray = value
        .map(item => (typeof item === 'string' ? item.trim() : item))
        .filter(item => item !== undefined && item !== null && item !== '');

      if (sanitizedArray.length > 0) {
        acc[key] = sanitizedArray;
      }

      return acc;
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed !== '') {
        acc[key] = trimmed;
      }
      return acc;
    }

    if (value !== undefined && value !== null) {
      acc[key] = value;
    }

    return acc;
  }, {});
}

function hasActiveFilters(filters) {
  const ignoredKeys = new Set(['page', 'pageSize', 'dataset']);
  const working = filters || {};

  return Object.keys(working).some(key => {
    if (ignoredKeys.has(key)) {
      return false;
    }

    const value = working[key];

    if (Array.isArray(value)) {
      return value.length > 0;
    }

    return value !== undefined && value !== null && value !== '';
  });
}

function cloneFilterObject(filters) {
  if (!filters || typeof filters !== 'object') {
    return {};
  }

  const clone = {};
  Object.keys(filters).forEach(key => {
    const value = filters[key];
    clone[key] = Array.isArray(value) ? value.slice() : value;
  });
  return clone;
}

function createFilterContext(filters, filtersAreSanitized = false) {
  const sanitizedFilters = filtersAreSanitized
    ? cloneFilterObject(filters || {})
    : sanitizeFilters(filters || {});

  const workingFilters = cloneFilterObject(sanitizedFilters);
  delete workingFilters.dataset;

  const lookups = {};
  const keysWithLookup = [
    'ano',
    'mes',
    'especialidade',
    'profissional',
    'situacao',
    'tipo',
    'tipoConsulta',
    'turno'
  ];

  keysWithLookup.forEach(key => {
    const rawValue = workingFilters[key];

    if (rawValue === undefined || rawValue === null || rawValue === '') {
      return;
    }

    if (Array.isArray(rawValue)) {
      if (rawValue.length === 0) {
        return;
      }

      const normalizedValues = rawValue
        .map(value => normalizeFilterValue(value))
        .filter(value => value !== '');

      if (normalizedValues.length > 0) {
        lookups[key] = new Set(normalizedValues);
      }

      return;
    }

    const normalized = normalizeFilterValue(rawValue);
    if (normalized !== '') {
      lookups[key] = new Set([normalized]);
    }
  });

  const searchTerm = workingFilters.search
    ? workingFilters.search.toString().toLowerCase()
    : '';

  return {
    filters: workingFilters,
    lookups: lookups,
    searchTerm: searchTerm,
    hasFilters: hasActiveFilters(workingFilters)
  };
}

function filterRecordsWithContext(data, context) {
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }

  if (!context || !context.hasFilters) {
    return data;
  }

  return data.filter(record => passesBasicFilters(record, context));
}

function registerCacheKey(prefix, key) {
  try {
    if (!prefix || !key) {
      return;
    }

    const props = PropertiesService.getScriptProperties();
    const rawRegistry = props.getProperty(CACHE_REGISTRY_KEY);
    const registry = rawRegistry ? JSON.parse(rawRegistry) : {};
    const normalizedPrefix = prefix;
    const existingKeys = Array.isArray(registry[normalizedPrefix]) ? registry[normalizedPrefix] : [];

    if (!existingKeys.includes(key)) {
      existingKeys.push(key);

      while (existingKeys.length > CACHE_REGISTRY_MAX_KEYS_PER_PREFIX) {
        const removedKey = existingKeys.shift();
        if (removedKey) {
          CacheService.getScriptCache().remove(removedKey);
        }
      }

      registry[normalizedPrefix] = existingKeys;
      props.setProperty(CACHE_REGISTRY_KEY, JSON.stringify(registry));
    }
  } catch (error) {
    console.error('Erro ao registrar chave de cache:', error);
  }
}

function clearCacheByPrefix(prefix) {
  try {
    if (!prefix) {
      return;
    }

    const props = PropertiesService.getScriptProperties();
    const rawRegistry = props.getProperty(CACHE_REGISTRY_KEY);
    if (!rawRegistry) {
      return;
    }

    const registry = JSON.parse(rawRegistry);
    const keys = registry[prefix];

    if (!Array.isArray(keys) || keys.length === 0) {
      delete registry[prefix];
      props.setProperty(CACHE_REGISTRY_KEY, JSON.stringify(registry));
      return;
    }

    const cache = CacheService.getScriptCache();
    const batchSize = 30;

    for (let i = 0; i < keys.length; i += batchSize) {
      const slice = keys.slice(i, i + batchSize);
      cache.removeAll(slice);
    }

    delete registry[prefix];
    props.setProperty(CACHE_REGISTRY_KEY, JSON.stringify(registry));
  } catch (error) {
    console.error('Erro ao limpar cache por prefixo:', error);
  }
}

function invalidateBaseDataCache(cacheInstance, datasetId) {
  const cache = cacheInstance || CacheService.getScriptCache();
  const datasetIds = datasetId ? [normalizeDatasetId(datasetId)] : Object.keys(CONFIG.datasets);

  datasetIds.forEach(id => {
    const metaKey = getBaseDataMetaKey(id);
    const metaRaw = cache.get(metaKey);

    if (!metaRaw) {
      return;
    }

    try {
      const meta = JSON.parse(metaRaw);
      const chunkKeys = Array.isArray(meta && meta.chunkKeys) ? meta.chunkKeys : [];
      const batchSize = 30;

      for (let i = 0; i < chunkKeys.length; i += batchSize) {
        const slice = chunkKeys.slice(i, i + batchSize);
        cache.removeAll(slice);
      }
    } catch (error) {
      console.error('Erro ao invalidar chunks do cache base:', error);
    } finally {
      cache.remove(metaKey);
    }
  });
}

function encodeCacheChunk(records) {
  const json = JSON.stringify(records || []);
  const compressed = Utilities.gzip(json);
  return Utilities.base64Encode(compressed.getBytes());
}

function decodeCacheChunk(encoded) {
  if (!encoded) {
    return null;
  }

  const bytes = Utilities.base64Decode(encoded);
  const blob = Utilities.newBlob(bytes);
  const decompressed = Utilities.ungzip(blob);
  const json = decompressed.getDataAsString('UTF-8');
  return JSON.parse(json);
}

function storeBaseDataInCache(cache, records, datasetId) {
  const normalizedDataset = normalizeDatasetId(datasetId);
  try {
    invalidateBaseDataCache(cache, normalizedDataset);

    const items = Array.isArray(records) ? records : [];

    if (items.length === 0) {
      cache.put(getBaseDataMetaKey(normalizedDataset), JSON.stringify({
        version: BASE_DATA_CACHE.version,
        dataset: normalizedDataset,
        chunkKeys: [],
        createdAt: new Date().toISOString()
      }), CONFIG.baseDataCacheDuration);
      return true;
    }

    const chunkKeys = [];
    const timestamp = new Date().getTime();
    const baseKey = `baseData_${normalizedDataset}_chunk_${timestamp}_`;
    let index = 0;
    let chunkIndex = 0;

    while (index < items.length) {
      let size = Math.min(CONFIG.baseDataCacheChunkSize, items.length - index);
      let encoded = '';
      let chunk = [];

      while (size > 0) {
        chunk = items.slice(index, index + size);
        encoded = encodeCacheChunk(chunk);

        if (encoded.length <= CONFIG.baseDataCacheMaxChunkBytes || size === 1) {
          break;
        }

        size = Math.floor(size / 2);
      }

      if (!encoded || encoded.length > CONFIG.baseDataCacheMaxChunkBytes) {
        console.warn('Não foi possível armazenar chunk da base de dados no cache (excedeu o limite).');
        invalidateBaseDataCache(cache, normalizedDataset);
        return false;
      }

      const chunkKey = `${baseKey}${chunkIndex}`;
      cache.put(chunkKey, encoded, CONFIG.baseDataCacheDuration);
      chunkKeys.push(chunkKey);
      index += chunk.length;
      chunkIndex++;
    }

    cache.put(getBaseDataMetaKey(normalizedDataset), JSON.stringify({
      version: BASE_DATA_CACHE.version,
      dataset: normalizedDataset,
      chunkKeys: chunkKeys,
      createdAt: new Date().toISOString()
    }), CONFIG.baseDataCacheDuration);

    return true;
  } catch (error) {
    console.error(`Erro ao armazenar base de dados no cache (${normalizedDataset}):`, error);
    invalidateBaseDataCache(cache, normalizedDataset);
    return false;
  }
}

function loadBaseDataFromCache(cache, datasetId) {
  const normalizedDataset = normalizeDatasetId(datasetId);
  const metaKey = getBaseDataMetaKey(normalizedDataset);
  const metaRaw = cache.get(metaKey);

  if (!metaRaw) {
    return null;
  }

  try {
    const meta = JSON.parse(metaRaw);

    if (!meta || meta.version !== BASE_DATA_CACHE.version) {
      return null;
    }

    const chunkKeys = Array.isArray(meta.chunkKeys) ? meta.chunkKeys : [];

    if (chunkKeys.length === 0) {
      return [];
    }

    const records = [];

    for (let i = 0; i < chunkKeys.length; i++) {
      const chunkKey = chunkKeys[i];
      const encoded = cache.get(chunkKey);

      if (!encoded) {
        return null;
      }

      const chunk = decodeCacheChunk(encoded);

      if (!Array.isArray(chunk)) {
        return null;
      }

      Array.prototype.push.apply(records, chunk);
    }

    return records;
  } catch (error) {
    console.error(`Erro ao carregar base de dados do cache (${normalizedDataset}):`, error);
    return null;
  }
}

function getCachedBaseData(spreadsheet, sheet, datasetId) {
  const normalizedDataset = normalizeDatasetId(datasetId);
  const cache = CacheService.getScriptCache();
  const cached = loadBaseDataFromCache(cache, normalizedDataset);

  if (cached !== null) {
    return cached;
  }

  const targetSheet = sheet || (spreadsheet ? getDatasetSheet(spreadsheet, normalizedDataset) : null);

  if (!targetSheet) {
    const definition = getDatasetDefinition(normalizedDataset);
    throw new Error(`Aba "${definition.sheetName}" não encontrada`);
  }

  const lastRow = targetSheet.getLastRow();

  if (lastRow <= 1) {
    storeBaseDataInCache(cache, [], normalizedDataset);
    return [];
  }

  const columnCount = getDatasetColumnCount(normalizedDataset, targetSheet);
  const dataRange = targetSheet.getRange(2, 1, lastRow - 1, columnCount);
  const data = dataRange.getValues();
  const records = data.map(row => mapDatasetRowToRecord(normalizedDataset, row));

  storeBaseDataInCache(cache, records, normalizedDataset);

  return records;
}

function paginateRecords(records, page, pageSize) {
  const list = Array.isArray(records) ? records : [];
  const size = Math.max(1, pageSize || CONFIG.pageSize);
  const currentPage = Math.max(1, Number(page) || 1);
  const startIndex = (currentPage - 1) * size;
  return list.slice(startIndex, startIndex + size);
}

function mapRecordToPageItem(record) {
  const item = coerceRowToRecord(record);

  return {
    prontuario: item.prontuario || '',
    paciente: item.paciente || '',
    profissional: item.profissional || '',
    tipo: item.tipo || '',
    especialidade: item.especialidade || '',
    situacao: item.situacao || '',
    ano: item.ano || '',
    mes: item.mes || '',
    data: formatDate(item.dataReferencia),
    quimioterapia: !!item.quimioterapia,
    cirurgia: !!item.cirurgia,
    turno: item.turno || ''
  };
}

function getTiposConsultaFromTemaSheet(temaSheet) {
  try {
    if (!temaSheet) {
      return { list: [], lookup: null };
    }

    const lastRow = temaSheet.getLastRow();
    if (lastRow <= 1) {
      return { list: [], lookup: null };
    }

    const values = temaSheet.getRange(2, 9, lastRow - 1, 1).getValues();
    const ordered = [];
    const lookup = Object.create(null);

    values.forEach(row => {
      const rawValue = row && row.length ? row[0] : '';
      if (!rawValue) {
        return;
      }

      const label = rawValue.toString().trim();
      if (!label) {
        return;
      }

      const key = label.toUpperCase();
      if (lookup[key]) {
        return;
      }

      lookup[key] = label;
      ordered.push(label);
    });

    return { list: ordered, lookup };
  } catch (error) {
    console.error('Erro ao obter tipos de consulta da aba tema:', error);
    return { list: [], lookup: null };
  }
}

// =============================================
// FUNÇÕES PRINCIPAIS - OTIMIZADAS
// =============================================

/**
 * OBTER OPÇÕES DE FILTRO (Aba "tema")
 * Busca apenas da aba tema, muito mais rápido
 */
function getFilterOptions(datasetId) {
  try {
    const normalizedDataset = normalizeDatasetId(datasetId);
    const cache = CacheService.getScriptCache();
    const cacheKey = `filterOptions_${normalizedDataset}`;
    const cached = cache.get(cacheKey);

    if (cached) {
      return JSON.parse(cached);
    }
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const temaSheet = getDatasetTemaSheet(spreadsheet, normalizedDataset);
    const temaSheetName = getDatasetThemeSheetName(normalizedDataset) || 'tema';
    const tiposConsultaInfo = getTiposConsultaFromTemaSheet(temaSheet);

    if (!temaSheet) {
      throw new Error(`Aba "${temaSheetName}" não encontrada`);
    }

    const data = temaSheet.getDataRange().getValues();

    // Mapeamento das colunas da aba tema conforme sua estrutura
    const options = {
      especialidades: getUniqueValuesFromColumn(data, 9),   // Coluna J
      profissionais: getUniqueValuesFromColumn(data, 6),    // Coluna G
      situacoes: getUniqueValuesFromColumn(data, 11),       // Coluna L
      anos: getUniqueValuesFromColumn(data, 14),            // Coluna O
      meses: getUniqueValuesFromColumn(data, 15),           // Coluna P
      tiposConsulta: tiposConsultaInfo.list.slice(),
      tipos: tiposConsultaInfo.list.slice(),
      categorias: getUniqueValuesFromColumn(data, 7),       // Coluna H
      comparecimentos: getUniqueValuesFromColumn(data, 17), // Coluna R
      turnos: getUniqueValuesFromColumn(data, 29)           // Coluna AD
    };
    
    // Limpar e ordenar opções
    Object.keys(options).forEach(key => {
      const values = Array.isArray(options[key]) ? options[key] : [];
      const cleaned = values
        .map(value => (value && value.toString ? value.toString().trim() : ''))
        .filter(value => value !== '');

      if (key === 'tiposConsulta' || key === 'tipos') {
        const seen = new Set();
        options[key] = cleaned.filter(value => {
          if (seen.has(value)) {
            return false;
          }
          seen.add(value);
          return true;
        });
      } else {
        options[key] = cleaned.sort();
      }
    });

    // Cache por 30 minutos (filtros mudam pouco)
    cache.put(cacheKey, JSON.stringify(options), 1800);
    registerCacheKey('filterOptions', cacheKey);

    return options;
  } catch (error) {
    console.error('Erro ao obter opções de filtro:', error);
    throw new Error('Não foi possível carregar as opções de filtro: ' + error.message);
  }
}

/**
 * OBTER DADOS PAGINADOS - PRINCIPAL FUNÇÃO OTIMIZADA
 * Processa dados em chunks para evitar limites de memória
 */
function getDashboardDataPaginated(filters = {}, page = 1) {
  const startTime = new Date().getTime();

  try {
    const sanitizedFilters = sanitizeFilters(filters || {});
    const datasetId = resolveDatasetIdFromFilters(sanitizedFilters);
    console.log('[Dashboard] Dataset selecionado (paginação):', datasetId);
    const filtersForProcessing = cloneFilterObject(sanitizedFilters);
    delete filtersForProcessing.dataset;
    const safePage = Math.max(1, Number(page) || 1);
    const cache = CacheService.getScriptCache();
    const sortedFilters = sortFilterObjectForCache(filtersForProcessing);
    const cacheKey = `dashboard_${datasetId}_${JSON.stringify(sortedFilters)}_${safePage}`;
    const cachedData = cache.get(cacheKey);

    if (cachedData) {
      console.log(`Cache hit para página ${safePage}`);
      return JSON.parse(cachedData);
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const databaseSheet = getDatasetSheet(spreadsheet, datasetId);

    if (!databaseSheet) {
      const definition = getDatasetDefinition(datasetId);
      throw new Error(`Aba "${definition.sheetName}" não encontrada`);
    }

    const baseData = getCachedBaseData(spreadsheet, databaseSheet, datasetId);
    const totalRecords = baseData.length;

    console.log(`Processando ${totalRecords} registros totais`);

    const filterContext = createFilterContext(filtersForProcessing, true);
    const workingData = filterRecordsWithContext(baseData, filterContext);
    const hasFilters = !!(filterContext && filterContext.hasFilters);
    const kpis = calculateKPIsWithFormulas(spreadsheet, filtersForProcessing, baseData, workingData, datasetId);

    const data = getPaginatedData(workingData, safePage);

    const result = {
      data: data,
      kpis: kpis,
      pagination: {
        currentPage: safePage,
        pageSize: CONFIG.pageSize,
        totalRecords: kpis.totalRegistros,
        totalPages: Math.ceil(Math.max(kpis.totalRegistros, 1) / CONFIG.pageSize)
      },
      performance: {
        processingTime: new Date().getTime() - startTime,
        totalRecords: totalRecords,
        hasFilters: hasFilters
      },
      timestamp: new Date().toISOString()
    };

    cache.put(cacheKey, JSON.stringify(result), 120);
    registerCacheKey('dashboard_', cacheKey);

    console.log(`Página ${safePage} processada em ${result.performance.processingTime}ms`);

    return result;
  } catch (error) {
    console.error('Erro em getDashboardDataPaginated:', error);
    throw new Error('Não foi possível carregar os dados: ' + error.message);
  }
}

/**
 * CALCULAR KPIs USANDO FÓRMULAS NATIVAS - SUPER RÁPIDO
 */
function calculateKPIsFromRecords(records) {
  if (!Array.isArray(records) || records.length === 0) {
    return getEmptyKPIs();
  }

  let totalConsultas = 0;
  let totalFaltas = 0;
  const pacientesSegmento = new Set();

  records.forEach(record => {
    const normalizedRecord = coerceRowToRecord(record);
    const situacao = normalizeSituacao(normalizedRecord.situacao);
    const prontuario = normalizedRecord.prontuario;

    if (situacao === 'CONCLUÍDO') {
      totalConsultas++;
      if (prontuario) {
        pacientesSegmento.add(prontuario);
      }
    }

    if (situacao === 'FALTOU') {
      totalFaltas++;
    }
  });

  const totalRegistros = records.length;
  const percentualAbsenteismo = totalRegistros > 0
    ? (totalFaltas / totalRegistros) * 100
    : 0;

  return {
    totalConsultas: totalConsultas,
    percentualAbsenteismo: percentualAbsenteismo,
    pacientesSegmento: pacientesSegmento.size,
    totalRegistros: totalRegistros,
    totalFaltas: totalFaltas
  };
}

function calculateKPIsWithFormulas(spreadsheet, filters = {}, baseData, preFilteredData, datasetId = DEFAULT_DATASET_ID) {
  const resolvedDataset = datasetId ? normalizeDatasetId(datasetId) : resolveDatasetIdFromFilters(filters);
  const normalizedDataset = resolvedDataset;
  console.log('[Dashboard] Dataset selecionado (KPIs):', normalizedDataset);

  try {
    const sanitizedFilters = sanitizeFilters(filters || {});
    delete sanitizedFilters.dataset;
    const cache = CacheService.getScriptCache();
    const sortedFilters = sortFilterObjectForCache(sanitizedFilters);
    const cacheKey = `kpis_${normalizedDataset}_${JSON.stringify(sortedFilters)}`;
    const cachedKPIs = cache.get(cacheKey);

    if (cachedKPIs) {
      return JSON.parse(cachedKPIs);
    }

    let workingData = Array.isArray(preFilteredData) ? preFilteredData : null;

    if (!workingData) {
      let data = Array.isArray(baseData) ? baseData : null;

      if (!data) {
        const databaseSheet = getDatasetSheet(spreadsheet, normalizedDataset);

        if (!databaseSheet) {
          const emptyKPIs = getEmptyKPIs();
          cache.put(cacheKey, JSON.stringify(emptyKPIs), 300);
          registerCacheKey('kpis_', cacheKey);
          return emptyKPIs;
        }

        data = getCachedBaseData(spreadsheet, databaseSheet, normalizedDataset);
      }

      if (!Array.isArray(data) || data.length === 0) {
        const emptyKPIs = getEmptyKPIs();
        cache.put(cacheKey, JSON.stringify(emptyKPIs), 300);
        registerCacheKey('kpis_', cacheKey);
        return emptyKPIs;
      }

      const context = createFilterContext(sanitizedFilters, true);
      workingData = filterRecordsWithContext(data, context);
    }

    const kpis = calculateKPIsFromRecords(workingData);

    cache.put(cacheKey, JSON.stringify(kpis), 300);
    registerCacheKey('kpis_', cacheKey);

    return kpis;
  } catch (error) {
    console.error(`Erro ao calcular KPIs com fórmulas (${normalizedDataset}):`, error);
    return getEmptyKPIs();
  }
}

/**
 * OBTER DADOS PAGINADOS A PARTIR DE REGISTROS PRONTOS
 */
function getPaginatedData(records, page) {
  const safePage = Math.max(1, Number(page) || 1);
  const pagedRecords = paginateRecords(records, safePage, CONFIG.pageSize);
  return pagedRecords.map(mapRecordToPageItem);
}

/**
 * APLICAR FILTROS AOS DADOS
 */
function applyFiltersToData(data, filters, filtersAreSanitized) {
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }

  const context = createFilterContext(filters, filtersAreSanitized);
  return filterRecordsWithContext(data, context);
}

// =============================================
// FUNÇÕES DE GRÁFICOS - OTIMIZADAS
// =============================================

/**
 * OBTER DADOS PARA GRÁFICOS (AGREGADOS)
 */
function getChartDataOptimized(filters = {}) {
  try {
    const sanitizedFilters = sanitizeFilters(filters || {});
    const datasetId = resolveDatasetIdFromFilters(sanitizedFilters);
    console.log('[Dashboard] Dataset selecionado (gráficos):', datasetId);
    const filtersForProcessing = cloneFilterObject(sanitizedFilters);
    delete filtersForProcessing.dataset;
    const cache = CacheService.getScriptCache();
    const sortedFilters = sortFilterObjectForCache(filtersForProcessing);
    const cacheKey = `charts_real_${datasetId}_${JSON.stringify(sortedFilters)}`;
    const cachedCharts = cache.get(cacheKey);

    if (cachedCharts) {
      return JSON.parse(cachedCharts);
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const databaseSheet = getDatasetSheet(spreadsheet, datasetId);
    const temaSheet = getDatasetTemaSheet(spreadsheet, datasetId);

    if (!databaseSheet) {
      return getEmptyChartData();
    }

    const baseData = getCachedBaseData(spreadsheet, databaseSheet, datasetId);

    if (!Array.isArray(baseData) || baseData.length === 0) {
      return getEmptyChartData();
    }

    const tiposConsultaInfo = getTiposConsultaFromTemaSheet(temaSheet);
    const chartData = buildRealChartData(baseData, filtersForProcessing, tiposConsultaInfo);

    cache.put(cacheKey, JSON.stringify(chartData), CONFIG.cacheDuration);
    registerCacheKey('charts_', cacheKey);

    return chartData;
  } catch (error) {
    console.error('Erro ao obter dados de gráficos:', error);
    return getEmptyChartData();
  }
}

/**
 * CONSTRUIR DADOS REAIS PARA GRÁFICOS
 */
function buildRealChartData(baseData, filters, tiposConsultaInfo) {
  const sanitizedFilters = sanitizeFilters(filters || {});
  const filterContext = createFilterContext(sanitizedFilters, true);
  const shouldFilter = !!(filterContext && filterContext.hasFilters);
  const especialidadeCount = {};
  const evolucaoCount = {};
  const situacoesCount = {};
  const faixaEtariaCount = {
    'Criança': 0,
    'Adulto': 0,
    'Idoso': 0
  };
  const generoCount = {};
  const tipoConsultaCount = {};
  const allowedSituacoes = new Set(['CONCLUÍDO', 'ANDAMENTO', 'ALTA AMBULATORIAL']);
  const tiposInfo = tiposConsultaInfo || {};
  const allowedTiposOrdered = Array.isArray(tiposInfo.list) ? tiposInfo.list : [];
  const allowedTiposLookup = tiposInfo.lookup || null;
  const dataSource = Array.isArray(baseData) ? baseData : [];

  dataSource.forEach(entry => {
    if (shouldFilter && !passesBasicFilters(entry, filterContext)) {
      return;
    }

    const record = coerceRowToRecord(entry);
    const situacao = normalizeSituacao(record.situacao);
    if (situacao) {
      situacoesCount[situacao] = (situacoesCount[situacao] || 0) + 1;
    }

    if (!allowedSituacoes.has(situacao)) {
      return;
    }

    const especialidade = normalizeFilterValue(record.especialidade);
    if (especialidade) {
      especialidadeCount[especialidade] = (especialidadeCount[especialidade] || 0) + 1;
    }

    const tipoConsultaValue = normalizeFilterValue(record.tipo);
    if (tipoConsultaValue) {
      const tipoKey = tipoConsultaValue.toUpperCase();
      if (!allowedTiposLookup || allowedTiposLookup[tipoKey]) {
        const displayLabel = allowedTiposLookup && allowedTiposLookup[tipoKey]
          ? allowedTiposLookup[tipoKey]
          : tipoConsultaValue;
        tipoConsultaCount[displayLabel] = (tipoConsultaCount[displayLabel] || 0) + 1;
      }
    }

    const dataReferencia = parseSheetDate(record.dataReferencia)
      || buildDateFromYearMonth(record.ano, record.mes);
    const faixa = getFaixaEtariaLabel(record.nascimento, dataReferencia);
    if (faixa) {
      faixaEtariaCount[faixa] = (faixaEtariaCount[faixa] || 0) + 1;
    }

    const genero = extractGenero(record);
    if (genero) {
      generoCount[genero] = (generoCount[genero] || 0) + 1;
    }

    const ano = normalizeFilterValue(record.ano);
    const mes = normalizeFilterValue(record.mes);

    if (ano && mes) {
      const chave = `${mes}/${ano}`;
      evolucaoCount[chave] = (evolucaoCount[chave] || 0) + 1;
    }
  });

  const especialidadeLabels = Object.keys(especialidadeCount)
    .sort((a, b) => (especialidadeCount[b] || 0) - (especialidadeCount[a] || 0));
  const especialidadeData = especialidadeLabels.map(label => especialidadeCount[label]);

  const evolucaoOrdenada = sortEvolucaoEntries(evolucaoCount);
  const faixaLabels = ['Criança', 'Adulto', 'Idoso'];
  const generoLabels = Object.keys(generoCount)
    .sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const desfechosLabels = Object.keys(situacoesCount);
  const desfechosData = desfechosLabels.map(label => situacoesCount[label]);
  const hasAllowedOrder = allowedTiposOrdered.length > 0;
  let tipoConsultaLabels = Object.keys(tipoConsultaCount);
  if (hasAllowedOrder) {
    tipoConsultaLabels = allowedTiposOrdered.filter(label => (tipoConsultaCount[label] || 0) > 0);
  } else {
    tipoConsultaLabels.sort((a, b) => (tipoConsultaCount[b] || 0) - (tipoConsultaCount[a] || 0));
  }
  const tipoConsultaData = tipoConsultaLabels.map(label => tipoConsultaCount[label] || 0);

  return {
    especialidadesValidas: {
      labels: especialidadeLabels,
      data: especialidadeData
    },
    evolucao: {
      labels: evolucaoOrdenada.labels,
      data: evolucaoOrdenada.data
    },
    situacoes: {
      labels: desfechosLabels,
      data: desfechosData
    },
    faixaEtaria: {
      labels: faixaLabels,
      data: faixaLabels.map(label => faixaEtariaCount[label] || 0)
    },
    genero: {
      labels: generoLabels,
      data: generoLabels.map(label => generoCount[label])
    },
    desfechos: {
      labels: desfechosLabels,
      data: desfechosData
    },
    tipoConsulta: {
      labels: tipoConsultaLabels,
      data: tipoConsultaData
    },
    especialidadesResumo: especialidadeLabels.map(label => ({
      especialidade: label,
      consultas: especialidadeCount[label]
    }))
  };
}

function sortEvolucaoEntries(evolucaoCount) {
  const entries = Object.keys(evolucaoCount || {}).map((key, index) => {
    const [mesRaw = '', anoRaw = ''] = key.split('/');
    const ano = parseInt(anoRaw.toString().replace(/\D/g, ''), 10);
    const mes = parseMesValue(mesRaw);

    return {
      key,
      value: evolucaoCount[key],
      ano: Number.isFinite(ano) ? ano : 0,
      mes: mes !== null ? mes : index + 1,
      index
    };
  });

  entries.sort((a, b) => {
    if (a.ano !== b.ano) return a.ano - b.ano;
    if (a.mes !== b.mes) return a.mes - b.mes;
    return a.index - b.index;
  });

  return {
    labels: entries.map(entry => entry.key),
    data: entries.map(entry => entry.value)
  };
}

function parseMesValue(value) {
  if (value === undefined || value === null) {
    return null;
  }

  const normalized = value.toString().trim().toLowerCase();

  if (!normalized) {
    return null;
  }

  const numeric = Number(normalized);
  if (Number.isFinite(numeric)) {
    return numeric;
  }

  const meses = {
    'janeiro': 1,
    'jan': 1,
    'fevereiro': 2,
    'fev': 2,
    'março': 3,
    'marco': 3,
    'mar': 3,
    'abril': 4,
    'abr': 4,
    'maio': 5,
    'mai': 5,
    'junho': 6,
    'jun': 6,
    'julho': 7,
    'jul': 7,
    'agosto': 8,
    'ago': 8,
    'setembro': 9,
    'set': 9,
    'outubro': 10,
    'out': 10,
    'novembro': 11,
    'nov': 11,
    'dezembro': 12,
    'dez': 12
  };

  return meses.hasOwnProperty(normalized) ? meses[normalized] : null;
}

function getFaixaEtariaLabel(nascimentoValue, referenciaValue) {
  const nascimento = parseSheetDate(nascimentoValue);
  if (!nascimento) {
    return '';
  }

  const referencia = parseSheetDate(referenciaValue) || new Date();
  const idade = calculateAgeInYears(nascimento, referencia);

  if (!Number.isFinite(idade) || idade < 0) {
    return '';
  }

  if (idade < 13) {
    return 'Criança';
  }

  if (idade < 60) {
    return 'Adulto';
  }

  return 'Idoso';
}

function calculateAgeInYears(nascimento, referencia) {
  if (!(nascimento instanceof Date) || !(referencia instanceof Date)) {
    return NaN;
  }

  let idade = referencia.getFullYear() - nascimento.getFullYear();
  const mes = referencia.getMonth() - nascimento.getMonth();

  if (mes < 0 || (mes === 0 && referencia.getDate() < nascimento.getDate())) {
    idade--;
  }

  return idade;
}

function parseSheetDate(value) {
  if (!value) {
    return null;
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && !isNaN(value)) {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return isNaN(date.getTime()) ? null : date;
  }

  if (typeof value === 'string') {
    const normalized = value.trim();
    if (!normalized) {
      return null;
    }

    const parts = normalized.split(/[\/-]/);
    if (parts.length === 3) {
      let [dia, mes, ano] = parts.map(part => part.replace(/\D/g, ''));

      if (ano && ano.length === 2) {
        ano = Number(ano) > 30 ? `19${ano}` : `20${ano}`;
      }

      const parsed = new Date(Number(ano), Number(mes) - 1, Number(dia));
      return isNaN(parsed.getTime()) ? null : parsed;
    }

    const parsedDate = new Date(normalized);
    return isNaN(parsedDate.getTime()) ? null : parsedDate;
  }

  return null;
}

function buildDateFromYearMonth(anoValue, mesValue) {
  if (!anoValue || !mesValue) {
    return null;
  }

  const ano = Number(anoValue.toString().replace(/\D/g, ''));
  const mesNumerico = parseMesValue(mesValue);

  if (!Number.isFinite(ano) || !Number.isFinite(mesNumerico) || mesNumerico < 1) {
    return null;
  }

  const data = new Date(ano, mesNumerico - 1, 1);
  return isNaN(data.getTime()) ? null : data;
}

function extractGenero(row) {
  if (!row) {
    return '';
  }

  if (Array.isArray(row)) {
    const candidateIndexes = [10, 12, 13];

    for (let i = 0; i < candidateIndexes.length; i++) {
      const index = candidateIndexes[i];
      if (typeof index !== 'number' || index >= row.length) {
        continue;
      }

      const genero = normalizeGenero(row[index]);
      if (genero) {
        return genero;
      }
    }

    return '';
  }

  const record = coerceRowToRecord(row);
  const candidates = [record.generoColJ, record.generoColL, record.generoColM];

  for (let i = 0; i < candidates.length; i++) {
    const genero = normalizeGenero(candidates[i]);
    if (genero) {
      return genero;
    }
  }

  return '';
}

function normalizeGenero(value) {
  if (!value) {
    return '';
  }

  const normalized = value.toString().trim().toUpperCase();
  if (!normalized) {
    return '';
  }

  if (['FEMININO', 'F', 'FEM'].includes(normalized)) {
    return 'Feminino';
  }

  if (['MASCULINO', 'M', 'MASC'].includes(normalized)) {
    return 'Masculino';
  }

  return '';
}

/**
 * VERIFICAR FILTROS BÁSICOS
 */
function passesBasicFilters(row, filters, filtersAreSanitized) {
  const context = filters && typeof filters === 'object' && filters.filters
    ? filters
    : createFilterContext(filters, filtersAreSanitized);

  const record = coerceRowToRecord(row);
  const safeFilters = context.filters || {};
  const lookups = context.lookups || {};

  if (!matchesFilterValue(record.ano, safeFilters.ano, lookups.ano)) return false;
  if (!matchesFilterValue(record.mes, safeFilters.mes, lookups.mes)) return false;
  if (!matchesFilterValue(record.especialidade, safeFilters.especialidade, lookups.especialidade)) return false;
  if (!matchesFilterValue(record.profissional, safeFilters.profissional, lookups.profissional)) return false;
  if (!matchesFilterValue(record.situacao, safeFilters.situacao, lookups.situacao)) return false;

  const tipoFilter = safeFilters.hasOwnProperty('tipoConsulta') && safeFilters.tipoConsulta !== undefined
    ? safeFilters.tipoConsulta
    : safeFilters.tipo;
  const tipoLookup = lookups.tipoConsulta || lookups.tipo;

  if (!matchesFilterValue(record.tipo, tipoFilter, tipoLookup)) return false;
  if (!matchesFilterValue(record.turno, safeFilters.turno, lookups.turno)) return false;

  const normalizedSearch = context.searchTerm || '';
  if (normalizedSearch) {
    const searchableFields = [
      record.prontuario,
      record.paciente,
      record.profissional,
      record.especialidade
    ];

    const matches = searchableFields.some(field =>
      field && field.toString().toLowerCase().includes(normalizedSearch)
    );

    if (!matches) {
      return false;
    }
  }

  return true;
}

function matchesFilterValue(itemValue, filterValue, lookup) {
  if (filterValue === undefined || filterValue === null || filterValue === '') {
    return true;
  }

  if (lookup instanceof Set) {
    if (lookup.size === 0) {
      return true;
    }

    const normalizedItem = normalizeFilterValue(itemValue);
    return lookup.has(normalizedItem);
  }

  if (Array.isArray(filterValue)) {
    if (filterValue.length === 0) {
      return true;
    }

    const normalizedItem = normalizeFilterValue(itemValue);
    return filterValue.some(value => normalizeFilterValue(value) === normalizedItem);
  }

  const normalizedItem = normalizeFilterValue(itemValue);
  const normalizedFilter = normalizeFilterValue(filterValue);

  return normalizedItem === normalizedFilter;
}

function normalizeFilterValue(value) {
  if (value === undefined || value === null) {
    return '';
  }

  if (typeof value === 'number') {
    return value.toString();
  }

  return value.toString().trim();
}

function normalizeSituacao(value) {
  return (value === undefined || value === null)
    ? ''
    : value.toString().trim().toUpperCase();
}

// =============================================
// FUNÇÕES DE EDIÇÃO E EXPORTAÇÃO
// =============================================

/**
 * ATUALIZAR SITUAÇÃO DE ATENDIMENTO
 */
function updateAtendimentoSituacao(prontuario, novaSituacao) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const databaseSheet = spreadsheet.getSheetByName(CONFIG.sheetNames.database);
    
    const lastRow = databaseSheet.getLastRow();
    const dataRange = databaseSheet.getRange(2, 1, lastRow - 1, 30);
    const data = dataRange.getValues();
    
    // Encontrar linha pelo prontuário (coluna C = índice 2)
    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === prontuario) {
        // Atualizar situação (coluna L = índice 11)
        databaseSheet.getRange(i + 2, 12).setValue(novaSituacao);
        
        // Invalidar cache relacionado
        clearRelevantCache();
        
        return {
          success: true,
          message: 'Situação atualizada com sucesso'
        };
      }
    }
    
    return {
      success: false,
      message: 'Prontuário não encontrado'
    };
  } catch (error) {
    console.error('Erro ao atualizar situação:', error);
    return {
      success: false,
      message: 'Erro ao atualizar situação: ' + error.message
    };
  }
}

/**
 * EXPORTAR DADOS PARA CSV
 */
function exportToCSV(filters = {}) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const datasetId = resolveDatasetIdFromFilters(filters);
    console.log('[Dashboard] Dataset selecionado (exportação CSV):', datasetId);
    const databaseSheet = getDatasetSheet(spreadsheet, datasetId);
    if (!databaseSheet) {
      return {
        success: false,
        message: 'Planilha de base de dados não encontrada'
      };
    }

    const baseData = getCachedBaseData(spreadsheet, databaseSheet, datasetId);

    if (!Array.isArray(baseData) || baseData.length === 0) {
      return {
        success: false,
        message: 'Nenhum dado para exportar'
      };
    }

    const sanitizedFilters = sanitizeFilters(filters || {});
    delete sanitizedFilters.dataset;
    const filteredData = applyFiltersToData(baseData, sanitizedFilters, true);

    if (filteredData.length === 0) {
      return {
        success: false,
        message: 'Nenhum dado encontrado com os filtros aplicados'
      };
    }
    
    // Criar CSV
    const headers = ['Prontuário', 'Paciente', 'Nascimento', 'Profissional', 'Tipo', 'Especialidade', 'Situação', 'Ano', 'Mês', 'Data', 'Quimioterapia', 'Cirurgia', 'Turno'];
    let csvContent = headers.join(';') + '\n';
    
    filteredData.forEach(row => {
      const record = coerceRowToRecord(row);
      const csvRow = [
        record.prontuario || '',
        `"${(record.paciente || '').replace(/"/g, '""')}"`,
        formatDate(record.nascimento),
        record.profissional || '',
        record.tipo || '',
        record.especialidade || '',
        record.situacao || '',
        record.ano || '',
        record.mes || '',
        formatDate(record.dataReferencia),
        record.quimioterapia ? 'SIM' : 'NÃO',
        record.cirurgia ? 'SIM' : 'NÃO',
        record.turno || ''
      ];

      csvContent += csvRow.join(';') + '\n';
    });
    
    // Criar arquivo no Google Drive
    const folder = DriveApp.getRootFolder();
    const fileName = `export_atendimentos_${new Date().toISOString().slice(0, 10)}.csv`;
    const file = folder.createFile(fileName, csvContent, 'text/csv');
    
    return {
      success: true,
      downloadUrl: file.getDownloadUrl(),
      fileName: fileName,
      recordCount: filteredData.length,
      message: `${filteredData.length} registros exportados com sucesso`
    };
  } catch (error) {
    console.error('Erro ao exportar CSV:', error);
    return {
      success: false,
      message: 'Erro ao exportar dados: ' + error.message
    };
  }
}

// =============================================
// FUNÇÕES AUXILIARES
// =============================================

/**
 * EXECUTAR FÓRMULA DE FORMA SEGURA
 */
function safeEvalFormula(sheet, formula) {
  try {
    const tempCell = sheet.getRange('Z1'); // Célula temporária
    tempCell.setFormula(formula);
    SpreadsheetApp.flush(); // Forçar atualização
    Utilities.sleep(100); // Pequena pausa
    const result = tempCell.getValue();
    tempCell.clear();
    return result || 0;
  } catch (error) {
    console.error('Erro na fórmula:', formula, error);
    return 0;
  }
}

/**
 * OBTER VALORES ÚNICOS DE UMA COLUNA
 */
function getUniqueValuesFromColumn(data, columnIndex) {
  const values = [];
  const seen = new Set();
  
  for (let i = 0; i < data.length; i++) {
    const value = data[i][columnIndex];
    if (value && !seen.has(value.toString().trim())) {
      seen.add(value.toString().trim());
      values.push(value.toString().trim());
    }
  }
  
  return values;
}

/**
 * FORMATAR DATA
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    if (typeof dateValue === 'object' && dateValue.getTime) {
      return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    
    const date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    
    return dateValue.toString();
  } catch (error) {
    return dateValue.toString();
  }
}

/**
 * LIMPAR CACHE RELEVANTE
 */
function clearRelevantCache() {
  const cache = CacheService.getScriptCache();
  invalidateBaseDataCache(cache);
  clearCacheByPrefix('dashboard_');
  clearCacheByPrefix('kpis_');
  clearCacheByPrefix('charts_');
  clearCacheByPrefix('filterOptions');
  cache.remove('filterOptions');

  console.log('Cache relevante invalidado');
}

/**
 * DADOS VAZIOS PARA KPIs
 */
function getEmptyKPIs() {
  return {
    totalConsultas: 0,
    percentualAbsenteismo: 0,
    pacientesSegmento: 0,
    totalRegistros: 0,
    totalFaltas: 0
  };
}

/**
 * DADOS VAZIOS PARA GRÁFICOS
 */
function getEmptyChartData() {
  return {
    especialidadesValidas: { labels: [], data: [] },
    evolucao: { labels: [], data: [] },
    situacoes: { labels: [], data: [] },
    faixaEtaria: { labels: [], data: [] },
    genero: { labels: [], data: [] },
    desfechos: { labels: [], data: [] },
    tipoConsulta: { labels: [], data: [] },
    especialidadesResumo: []
  };
}

// =============================================
// FUNÇÕES DE TESTE E DEBUG
// =============================================

/**
 * TESTAR CONEXÃO E PERFORMANCE
 */
function testConnection() {
  try {
    const startTime = new Date().getTime();
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const databaseSheet = spreadsheet.getSheetByName(CONFIG.sheetNames.database);
    const temaSheet = getDatasetTemaSheet(spreadsheet, DEFAULT_DATASET_ID);
    
    if (!databaseSheet || !temaSheet) {
      throw new Error('Abas não encontradas');
    }
    
    const lastRow = databaseSheet.getLastRow();
    const lastColumn = databaseSheet.getLastColumn();
    
    const filterOptions = getFilterOptions(DEFAULT_DATASET_ID);
    const kpis = calculateKPIsWithFormulas(spreadsheet, {}, null, null, DEFAULT_DATASET_ID);
    
    const processingTime = new Date().getTime() - startTime;
    
    return {
      success: true,
      message: 'Conexão estabelecida com sucesso',
      performance: {
        processingTime: processingTime + 'ms',
        totalRecords: lastRow - 1,
        totalColumns: lastColumn,
        cacheEnabled: true
      },
      sheets: {
        database: CONFIG.sheetNames.database,
        tema: getDatasetThemeSheetName(DEFAULT_DATASET_ID),
        status: 'OK'
      },
      kpis: kpis,
      filterOptions: {
        especialidades: filterOptions.especialidades.length,
        profissionais: filterOptions.profissionais.length,
        anos: filterOptions.anos.length
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'Erro: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * LIMPAR TODO O CACHE
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    invalidateBaseDataCache(cache);
    clearCacheByPrefix('dashboard_');
    clearCacheByPrefix('kpis_');
    clearCacheByPrefix('charts_');
    clearCacheByPrefix('filterOptions');
    cache.remove('filterOptions');
    return {
      success: true,
      message: 'Cache limpo com sucesso'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao limpar cache: ' + error.message
    };
  }
}

// =============================================
// COMPATIBILIDADE - FUNÇÕES LEGACY
// =============================================

/**
 * FUNÇÃO LEGACY - REDIRECIONA PARA NOVA FUNÇÃO PAGINADA
 */
function getDashboardData(filters = {}) {
  console.log('Usando função legacy getDashboardData - redirecionando para paginada');
  return getDashboardDataPaginated(filters, 1);
}

/**
 * FUNÇÃO LEGACY PARA GRÁFICOS
 */
function getChartData(filters = {}) {
  console.log('Usando função legacy getChartData - redirecionando para otimizada');
  return getChartDataOptimized(filters);
}

// =============================================
// BI CIRÚRGICO - FUNÇÕES DEDICADAS
// =============================================

function getCirurgicoBaseData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'cirurgico_base_v1';
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      console.warn('Falha ao desserializar cache do BI Cirúrgico:', error);
    }
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = getDatasetSheet(spreadsheet, 'cirurgico');

    if (!sheet) {
      throw new Error('Planilha do BI Cirúrgico não encontrada');
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow <= 1 || lastColumn <= 0) {
      const emptyPayload = {
        success: true,
        data: [],
        lastUpdated: new Date().toISOString()
      };
      cache.put(cacheKey, JSON.stringify(emptyPayload), CONFIG.cacheDuration || 600);
      registerCacheKey('cirurgico_base', cacheKey);
      return emptyPayload;
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(value => (value && value.toString ? value.toString().trim() : ''));

    const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    const data = values.map(row => {
      const record = {};
      headers.forEach((header, index) => {
        if (header) {
          record[header] = row[index];
        }
      });
      Object.assign(record, mapCirurgicoRowToRecord(row));
      enrichCirurgicoRecord(record);
      return record;
    });

    const payload = {
      success: true,
      data,
      lastUpdated: new Date().toISOString()
    };

    cache.put(cacheKey, JSON.stringify(payload), CONFIG.cacheDuration || 600);
    registerCacheKey('cirurgico_base', cacheKey);

    return payload;
  } catch (error) {
    console.error('Erro ao carregar base do BI Cirúrgico:', error);
    return {
      success: false,
      error: error.message,
      data: [],
      lastUpdated: new Date().toISOString()
    };
  }
}

function normalizeHeaderKey(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return value
    .toString()
    .trim()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toUpperCase()
    .replace(/[\s_/\\-]+/g, '_')
    .replace(/[^A-Z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

function enrichCirurgicoRecord(record) {
  if (!record || typeof record !== 'object') {
    return record;
  }

  Object.keys(record).forEach(key => {
    if (!key) {
      return;
    }

    const underscoredKey = key.replace(/\s+/g, '_');
    if (underscoredKey && underscoredKey !== key && record[underscoredKey] === undefined) {
      record[underscoredKey] = record[key];
    }

    const normalizedKey = normalizeHeaderKey(key);
    if (normalizedKey && record[normalizedKey] === undefined) {
      record[normalizedKey] = record[key];
    }
  });

  return record;
}

function getCirurgicoField(record, keys) {
  if (!record) {
    return '';
  }

  const normalizedKeys = Array.isArray(keys) ? keys : [keys];
  for (let i = 0; i < normalizedKeys.length; i += 1) {
    const key = normalizedKeys[i];
    if (key && Object.prototype.hasOwnProperty.call(record, key)) {
      return record[key];
    }
  }

  const normalizedRecordMap = {};
  Object.keys(record).forEach(existingKey => {
    const normalizedKey = normalizeHeaderKey(existingKey);
    if (normalizedKey && normalizedRecordMap[normalizedKey] === undefined) {
      normalizedRecordMap[normalizedKey] = existingKey;
    }
  });

  for (let i = 0; i < normalizedKeys.length; i += 1) {
    const key = normalizedKeys[i];
    if (!key) {
      continue;
    }

    const normalizedKey = normalizeHeaderKey(key);
    const resolvedKey = normalizedRecordMap[normalizedKey];
    if (resolvedKey && Object.prototype.hasOwnProperty.call(record, resolvedKey)) {
      return record[resolvedKey];
    }
  }

  return '';
}

function normalizeSheetString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim();
}

function filterCirurgicoRecords(filters = {}) {
  const base = getCirurgicoBaseData();

  if (!base.success) {
    return base;
  }

  const { data } = base;
  const normalizedFilters = filters || {};
  const searchTerm = normalizeSheetString(normalizedFilters.search).toLowerCase();
  const programaFiltro = normalizeSheetString(normalizedFilters.programa).toLowerCase();
  const anoFiltro = normalizeSheetString(normalizedFilters.ano).toLowerCase();
  const mesFiltro = normalizeSheetString(normalizedFilters.mes).toLowerCase();
  const especialidadeFiltro = normalizeSheetString(normalizedFilters.especialidade).toLowerCase();
  const situacaoFiltro = normalizeSheetString(normalizedFilters.situacao).toLowerCase();

  const filtered = data.filter(record => {
    const programa = normalizeSheetString(getCirurgicoField(record, ['PROGRAMA'])).toLowerCase();
    if (programaFiltro && programa !== programaFiltro) {
      return false;
    }

    const ano = normalizeSheetString(getCirurgicoField(record, ['ANO', 'Ano'])).toLowerCase();
    if (anoFiltro && ano !== anoFiltro) {
      return false;
    }

    const mes = normalizeSheetString(getCirurgicoField(record, ['MÊS', 'MES'])).toLowerCase();
    if (mesFiltro && mes !== mesFiltro) {
      return false;
    }

    const especialidade = normalizeSheetString(getCirurgicoField(record, ['ESPECIALIDADE'])).toLowerCase();
    if (especialidadeFiltro && especialidade !== especialidadeFiltro) {
      return false;
    }

    const situacao = normalizeSheetString(getCirurgicoField(record, ['SITUAÇÃO_FASTMEDIC', 'SITUACAO_FASTMEDIC'])).toLowerCase();
    if (situacaoFiltro && situacao !== situacaoFiltro) {
      return false;
    }

    if (!searchTerm) {
      return true;
    }

    const prontuario = normalizeSheetString(getCirurgicoField(record, ['PRONTUÁRIO', 'PRONTUARIO', 'PRONT'])).toLowerCase();
    const nome = normalizeSheetString(getCirurgicoField(record, ['NOME'])).toLowerCase();
    const municipio = normalizeSheetString(getCirurgicoField(record, ['MUNICÍPIO', 'MUNICIPIO'])).toLowerCase();
    const procedimento = normalizeSheetString(getCirurgicoField(record, ['PROCEDIMENTO', 'PROCEDIMENTO INDICADO'])).toLowerCase();

    return (
      prontuario.includes(searchTerm)
      || nome.includes(searchTerm)
      || municipio.includes(searchTerm)
      || procedimento.includes(searchTerm)
    );
  });

  return {
    success: true,
    data: filtered,
    lastUpdated: base.lastUpdated
  };
}

function buildCirurgicoMetrics(data) {
  const total = data.length;
  let surgeriesCount = 0;
  let pendingCount = 0;
  let federalCount = 0;
  let stateCount = 0;
  let confirmedCount = 0;
  const waitTimes = [];

  data.forEach(record => {
    const dataCirurgia = getCirurgicoField(record, ['DATA_DA_CIRURGIA']);
    if (dataCirurgia) {
      surgeriesCount += 1;
    }

    const status = normalizeSheetString(getCirurgicoField(record, ['STATUS_DO_AGENDAMENTO'])).toUpperCase();
    if (status === 'PENDENTE' || status === 'AGENDADA') {
      pendingCount += 1;
    }

    const programa = normalizeSheetString(getCirurgicoField(record, ['PROGRAMA'])).toUpperCase();
    if (programa === 'FEDERAL') {
      federalCount += 1;
    }
    if (programa === 'ESTADUAL') {
      stateCount += 1;
    }

    const confirmacao = normalizeSheetString(getCirurgicoField(record, ['CONFIRMAÇÃO_DA_INTERNAÇÃO', 'CONFIRMACAO_DA_INTERNAÇÃO', 'CONFIRMACAO_DA_INTERNACAO'])).toUpperCase();
    if (confirmacao === 'CONFIRMADA') {
      confirmedCount += 1;
    }

    const tempoEspera = getCirurgicoField(record, ['TEMPO_DE_ESPERA', 'TEMPO DE ESPERA PARA AGENDAMENTO CIRURGICO']);
    const numericTime = Number(tempoEspera);
    if (Number.isFinite(numericTime) && numericTime > 0) {
      waitTimes.push(numericTime);
    }
  });

  const confirmationRate = total > 0 ? Number((confirmedCount / total) * 100) : 0;
  const avgWaitTime = waitTimes.length
    ? Number(waitTimes.reduce((sum, value) => sum + value, 0) / waitTimes.length)
    : 0;

  return {
    surgeriesCount,
    pendingCount,
    federalCount,
    stateCount,
    confirmationRate,
    avgWaitTime,
    totalRecords: total
  };
}

const CIRURGICO_MONTH_LABELS = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

const CIRURGICO_MONTH_ORDER = CIRURGICO_MONTH_LABELS.reduce((acc, label, index) => {
  const normalized = label.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();
  acc[normalized] = index + 1;
  return acc;
}, {});

function normalizeCirurgicoMonth(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const raw = value.toString().trim();
  if (!raw) {
    return '';
  }

  const numeric = Number(raw);
  if (Number.isFinite(numeric)) {
    const index = Math.min(Math.max(1, Math.floor(numeric)), 12) - 1;
    return CIRURGICO_MONTH_LABELS[index];
  }

  const normalized = raw.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();
  const order = CIRURGICO_MONTH_ORDER[normalized];
  if (order) {
    return CIRURGICO_MONTH_LABELS[order - 1];
  }
  return raw;
}

function sortCirurgicoMonthlyLabels(labels) {
  return labels.sort((a, b) => {
    const [monthA, yearA] = splitMonthYear(a);
    const [monthB, yearB] = splitMonthYear(b);

    const numericYearA = Number(yearA);
    const numericYearB = Number(yearB);

    if (Number.isFinite(numericYearA) && Number.isFinite(numericYearB) && numericYearA !== numericYearB) {
      return numericYearA - numericYearB;
    }

    const normalizedMonthA = monthA.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();
    const normalizedMonthB = monthB.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();

    const orderA = CIRURGICO_MONTH_ORDER[normalizedMonthA] || 99;
    const orderB = CIRURGICO_MONTH_ORDER[normalizedMonthB] || 99;

    return orderA - orderB;
  });
}

function splitMonthYear(label) {
  if (!label) {
    return ['', ''];
  }
  const parts = label.toString().trim().split(' ');
  if (parts.length >= 2) {
    const year = parts.pop();
    const month = parts.join(' ');
    return [month, year];
  }
  return [label, ''];
}

function buildCirurgicoCharts(data) {
  const specialtyCounts = {};
  const monthlyCounts = {};
  const statusCounts = {};

  data.forEach(record => {
    const specialty = normalizeSheetString(getCirurgicoField(record, ['ESPECIALIDADE'])) || 'Não informado';
    specialtyCounts[specialty] = (specialtyCounts[specialty] || 0) + 1;

    const rawMonth = getCirurgicoField(record, ['MÊS', 'MES']);
    const rawYear = getCirurgicoField(record, ['ANO', 'Ano']);
    const monthLabel = normalizeCirurgicoMonth(rawMonth);
    if (monthLabel) {
      const label = rawYear ? `${monthLabel} ${rawYear}` : monthLabel;
      monthlyCounts[label] = (monthlyCounts[label] || 0) + 1;
    }

    const status = normalizeSheetString(getCirurgicoField(record, ['STATUS_DO_AGENDAMENTO'])) || 'Não informado';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  const sortedMonthlyLabels = sortCirurgicoMonthlyLabels(Object.keys(monthlyCounts));

  return {
    specialty: {
      labels: Object.keys(specialtyCounts),
      data: Object.values(specialtyCounts)
    },
    monthly: {
      labels: sortedMonthlyLabels,
      data: sortedMonthlyLabels.map(label => monthlyCounts[label])
    },
    status: {
      labels: Object.keys(statusCounts),
      data: Object.values(statusCounts)
    }
  };
}

function buildCirurgicoInsights(data, metrics) {
  const insights = [];
  insights.push({
    icon: 'fas fa-procedures',
    text: `${metrics.surgeriesCount || 0} cirurgias realizadas no período.`
  });
  insights.push({
    icon: 'fas fa-calendar-check',
    text: `${metrics.pendingCount || 0} agendamentos aguardando execução.`
  });
  insights.push({
    icon: 'fas fa-chart-pie',
    text: `${metrics.federalCount || 0} pacientes federais e ${metrics.stateCount || 0} estaduais.`
  });
  if (Number.isFinite(metrics.avgWaitTime) && metrics.avgWaitTime > 0) {
    insights.push({
      icon: 'fas fa-hourglass-half',
      text: `Tempo médio de espera de ${Math.round(metrics.avgWaitTime)} dias.`
    });
  }
  return insights;
}

function getCirurgicoFilterOptions() {
  const base = getCirurgicoBaseData();
  if (!base.success) {
    return {
      success: false,
      error: base.error || 'Não foi possível carregar os dados cirúrgicos',
      filters: {}
    };
  }

  const unique = {
    programs: new Set(),
    years: new Set(),
    months: new Set(),
    specialties: new Set(),
    situations: new Set()
  };

  base.data.forEach(record => {
    const programa = normalizeSheetString(getCirurgicoField(record, ['PROGRAMA']));
    if (programa) {
      unique.programs.add(programa);
    }

    const ano = normalizeSheetString(getCirurgicoField(record, ['ANO', 'Ano']));
    if (ano) {
      unique.years.add(ano);
    }

    const mes = normalizeSheetString(getCirurgicoField(record, ['MÊS', 'MES']));
    if (mes) {
      unique.months.add(mes);
    }

    const especialidade = normalizeSheetString(getCirurgicoField(record, ['ESPECIALIDADE']));
    if (especialidade) {
      unique.specialties.add(especialidade);
    }

    const situacao = normalizeSheetString(getCirurgicoField(record, ['SITUAÇÃO_FASTMEDIC', 'SITUACAO_FASTMEDIC']));
    if (situacao) {
      unique.situations.add(situacao);
    }
  });

  const monthsOrdered = Array.from(unique.months).sort((a, b) => {
    const normalizedA = normalizeCirurgicoMonth(a);
    const normalizedB = normalizeCirurgicoMonth(b);
    const keyA = normalizedA.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();
    const keyB = normalizedB.normalize('NFD').replace(/\p{Diacritic}/gu, '').toUpperCase();
    const orderA = CIRURGICO_MONTH_ORDER[keyA] || 99;
    const orderB = CIRURGICO_MONTH_ORDER[keyB] || 99;
    return orderA - orderB;
  });

  return {
    success: true,
    filters: {
      programs: Array.from(unique.programs).sort(),
      years: Array.from(unique.years).sort((a, b) => Number(b) - Number(a)),
      months: monthsOrdered,
      specialties: Array.from(unique.specialties).sort(),
      situations: Array.from(unique.situations).sort()
    }
  };
}

function getCirurgicoDashboardSnapshot(filters = {}) {
  const filtered = filterCirurgicoRecords(filters);
  if (!filtered.success) {
    return filtered;
  }

  const metrics = buildCirurgicoMetrics(filtered.data);
  const charts = buildCirurgicoCharts(filtered.data);
  const insights = buildCirurgicoInsights(filtered.data, metrics);

  return {
    success: true,
    metrics,
    charts,
    table: {
      rows: filtered.data.slice(0, 150),
      total: filtered.data.length
    },
    insights,
    performance: {
      totalRecords: filtered.data.length
    },
    lastUpdated: filtered.lastUpdated
  };
}

function exportCirurgicoData(filters = {}) {
  try {
    const filtered = filterCirurgicoRecords(filters);
    if (!filtered.success) {
      return filtered;
    }

    return {
      success: true,
      message: `Exportação preparada para ${filtered.data.length} registros cirúrgicos`,
      total: filtered.data.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

function getCirurgicoDataFromSheet() {
  return getCirurgicoBaseData();
}

function getCirurgicoFilteredData(filters = {}) {
  const filtered = filterCirurgicoRecords(filters);
  if (!filtered.success) {
    return filtered;
  }

  return {
    success: true,
    data: filtered.data,
    total: filtered.data.length,
    lastUpdated: filtered.lastUpdated || new Date().toISOString()
  };
}

function getCirurgicoKPIMetrics(filters = {}) {
  const filtered = filterCirurgicoRecords(filters);
  if (!filtered.success) {
    return filtered;
  }

  return {
    success: true,
    metrics: buildCirurgicoMetrics(filtered.data)
  };
}

function getCirurgicoChartData(filters = {}) {
  const filtered = filterCirurgicoRecords(filters);
  if (!filtered.success) {
    return filtered;
  }

  const charts = buildCirurgicoCharts(filtered.data);
  const metrics = buildCirurgicoMetrics(filtered.data);
  const insights = buildCirurgicoInsights(filtered.data, metrics);

  return {
    success: true,
    charts,
    insights
  };
}
