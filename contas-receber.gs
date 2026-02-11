/**
 * ============ CONTAS A RECEBER (Bulk Data /income) ============
 * - Uma única aba (usa o nome existente, ignorando maiúsculas/minúsculas)
 * - Traz todas as colunas escalares da PARCELA + do RECIBO (1 linha por parcela×recibo)
 * - EXCLUI: receiptsCategories, bankMovements e financialCategories
 * - Filtro só do CC 78 (no servidor)
 * - Janelas de 92 dias (2023-01-01..2030-01-01 | selectionType=D)
 * - Ajusta automaticamente a grade (linhas/colunas) antes de escrever (corrige "exceeds grid limits")
 * - Cabeçalhos traduzidos para PT-BR
 */

// ========== CONFIG ==========
const CONFIG_INCOME = {
  SIENGE_SUBDOMAIN: PropertiesService.getScriptProperties().getProperty('SIENGE_SUBDOMAIN') || 'm21',
  SIENGE_USER:      PropertiesService.getScriptProperties().getProperty('SIENGE_USER'),
  SIENGE_PASS:      PropertiesService.getScriptProperties().getProperty('SIENGE_PASS'),

  START_DATE: '2023-01-01',
  END_DATE:   '2030-01-01',
  SELECTION_TYPE: 'D',            // I=Emissão, D=Vencimento, P=Pagamento, B=Competência
  COMPANY_ID: 2007,               // sua empresa
  COST_CENTER_FILTER: 78,         // SOMENTE CC 78

  WINDOW_DAYS: 92,
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),

  // Nome de referência da aba (o script encontra a existente ignorando caixa)
  SHEET_NAME: 'Contas a Receber',

  // --- Exclusões solicitadas (informadas pelo usuário) ---
  EXCLUDE_KEYS: [
    'bearerId',
    'businessAreaId', 'businessAreaName',
    'businessTypeId', 'businessTypeName',
    'companyId', 'companyName',
    'embeddedInterestAmount',
    'groupCompanyId', 'groupCompanyName',
    'holdingId', 'holdingName',
    'interestRate',
    'projectId',
    'receipt.taxAmount',
    'subJudicie',
    'subsidiaryId', 'subsidiaryName',
    'taxAmount'
  ],

  EXCLUDE_PREFIXES: [
    // vazio
  ],

  // Encolher grade exatamente ao que foi escrito (apaga colunas removidas)
  SHRINK_SHEET: true
};

// ========== ENTRADA ==========
function atualizarContasAReceber() {
  console.log('🚀 Iniciando Contas a Receber (aba única, PT-BR, CC 78)…');

  const bruto = fetchIncomeChunked_();
  console.log(`📊 Parcelas retornadas: ${bruto.length}`);

  const flat = buildFlatRows_(bruto); // { headers (raw), rows }

  // aplica exclusões por chave/prefixo (antes da tradução)
  const filtered = applyColumnExclusions_(flat.headers, flat.rows, CONFIG_INCOME.EXCLUDE_KEYS || [], CONFIG_INCOME.EXCLUDE_PREFIXES || []);
  const headersRaw = filtered.headers;
  const rows       = filtered.rows;

  const headersPt = headersRaw.map(translateKeyPt_); // nomes PT-BR

  escreverNaPlanilhaIncomeFlat_(headersPt, headersRaw, rows);

  console.log(`✅ Concluído: ${rows.length} linhas escritas.`);
}

// ========== FETCH EM JANELAS ==========
function fetchIncomeChunked_() {
  let cursor = new Date(CONFIG_INCOME.START_DATE);
  const end  = new Date(CONFIG_INCOME.END_DATE);
  const all = [];
  let janela = 0;

  while (cursor <= end) {
    janela++;
    const from = new Date(cursor);
    const to   = new Date(cursor);
    to.setDate(to.getDate() + (CONFIG_INCOME.WINDOW_DAYS - 1));
    if (to > end) to.setTime(end.getTime());

    const startStr = Utilities.formatDate(from, 'GMT-3', 'yyyy-MM-dd');
    const endStr   = Utilities.formatDate(to,   'GMT-3', 'yyyy-MM-dd');

    const chunk = fetchIncomeOnce_(startStr, endStr);
    console.log(`📦 Janela #${janela} ${startStr}..${endStr}: ${chunk.length} parcelas`);
    Array.prototype.push.apply(all, chunk);

    to.setDate(to.getDate() + 1); // avança para dia seguinte
    cursor = to;
  }
  return all;
}

function fetchIncomeOnce_(startDate, endDate) {
  const base = `https://api.sienge.com.br/${CONFIG_INCOME.SIENGE_SUBDOMAIN}/public/api/bulk-data/v1/income`;
  const q = [];
  q.push('startDate=' + encodeURIComponent(startDate));
  q.push('endDate='   + encodeURIComponent(endDate));
  q.push('selectionType=' + encodeURIComponent(CONFIG_INCOME.SELECTION_TYPE));
  q.push('costCentersId=' + encodeURIComponent(String(CONFIG_INCOME.COST_CENTER_FILTER))); // SOMENTE CC 78
  if (CONFIG_INCOME.COMPANY_ID) q.push('companyId=' + encodeURIComponent(String(CONFIG_INCOME.COMPANY_ID)));
  const url = `${base}?${q.join('&')}`;

  const options = {
    method: 'get',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(`${CONFIG_INCOME.SIENGE_USER}:${CONFIG_INCOME.SIENGE_PASS}`),
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  };

  console.log('📡 GET:', url);
  const res   = UrlFetchApp.fetch(url, options);
  const code  = res.getResponseCode();
  const ctype = (res.getHeaders()['Content-Type'] || '').toLowerCase();

  if (code !== 200) throw new Error(`HTTP ${code} :: ${res.getContentText()}`);
  if (!ctype.includes('application/json')) {
    throw new Error(`Conteúdo não-JSON em ${startDate}..${endDate} (Content-Type: ${ctype}). Reduza WINDOW_DAYS.`);
  }

  const json = JSON.parse(res.getContentText());
  return Array.isArray(json && json.data) ? json.data : [];
}

// ========== FLATTEN (Parcela + Receipts, excluindo arrays indesejados) ==========
// Esta versão extrai "Número da Unidade" usando prioritariamente a coluna de unidades.
// Só usa documentNumber quando unidade estiver em branco E o tipo do documento for ADIANTAMENTO...
// Depois disso, faz um pós-processamento: preenche linhas vazias repetindo o número já conhecido para o mesmo cliente.
function buildFlatRows_(data) {
  const headersSet = new Set();
  const rows = [];

  function isScalar(v) { return (v === null) || (typeof v !== 'object'); }
  function copyScalarsAndOneLevel(obj, prefix, target, headers) {
    Object.keys(obj || {}).forEach(k => {
      const key = prefix ? `${prefix}.${k}` : k;
      const val = obj[k];
      if (isScalar(val)) {
        target[key] = val;
        headers.add(key);
      } else if (!Array.isArray(val) && val) {
        Object.keys(val).forEach(sub => {
          const sk = `${key}.${sub}`;
          const sv = val[sub];
          if (isScalar(sv)) {
            target[sk] = sv;
            headers.add(sk);
          }
        });
      }
    });
  }

  /* -------------------- Helpers locais -------------------- */

  // todas sequências de dígitos numa string
  function _allDigitsFromString(s) {
    if (!s && s !== 0) return [];
    const arr = String(s).match(/\d+/g);
    return arr ? arr.map(x => x.trim()) : [];
  }

  // padrões preferenciais (APTO/ADCAPTO/CTAPTO/ADIANTAPTO/CTP/PENTHOUSE etc.)
  function _preferPatternNumber(s) {
    if (!s) return '';
    const str = String(s).trim();
    const patterns = [
      /(?:APTO|APT|ADCAPTO|CTAPTO|ADIANTAPTO)\D*?(\d{1,6})/i,
      /(?:CTP|CT|AD)\D*?(\d{1,6})$/i,
      /(PENTHOUSE|PENT)\D*?(\d{1,6})/i,
      /[A-Za-z]+\D*?(\d{1,6})/i
    ];
    for (const re of patterns) {
      const m = str.match(re);
      if (m) {
        for (let i = m.length-1; i>=1; i--) {
          if (/\d+/.test(m[i])) return m[i];
        }
      }
    }
    return '';
  }

  // extrai candidatos numa string (ordem preservada)
  function _extractCandidates(s) {
    if (!s && s !== 0) return [];
    const parts = String(s).split(/\s*\|\s*|\n|;|,/).map(p => p.trim()).filter(Boolean);
    const out = [];
    for (const p of parts) {
      const byPat = _preferPatternNumber(p);
      if (byPat) out.push(byPat);
      else {
        _allDigitsFromString(p).forEach(d => out.push(d));
      }
    }
    return [...new Set(out)];
  }

  // recupera nome/tipo do documento (tolerante)
  function _getDocumentTypeName(item, baseRow) {
    const candidates = [];
    if (item) {
      if (item.documentIdentificationName) candidates.push(String(item.documentIdentificationName));
      if (item.documentName) candidates.push(String(item.documentName));
      if (item.document && item.document.name) candidates.push(String(item.document.name));
      if (item.documentIdentification && item.documentIdentification.name) candidates.push(String(item.documentIdentification.name));
      if (item.documentTypeName) candidates.push(String(item.documentTypeName));
    }
    if (baseRow && baseRow['documentIdentificationName']) candidates.push(String(baseRow['documentIdentificationName']));
    return candidates.map(c=>c.trim()).filter(Boolean).join(' | ');
  }

  // validação por tamanho quando extraindo de documentNumber
  const MIN_ACCEPT_LEN = 2; // evita aceitar '2' aleatório
  const MAX_ACCEPT_LEN = 6; // evita CPFs/IDs longos

  function _isAdiantamentoType(docTypeName) {
    if (!docTypeName) return false;
    return /ADIANTAMENTO\s+DE\s+RECEBIMENTO\s+DE\s+CLIENTES/i.test(docTypeName) ||
           /ADIANTAMENTO/i.test(docTypeName);
  }

  // extrai de campos de unidade explícitos (mainUnit, unit, unitName, units[], etc.)
  function _extractFromUnitsFields(item, baseRow) {
    const candidates = [];
    const fields = [
      'mainUnit','unit','unitName','unitNumber','unitId','unitCode',
      'unity','apartment','apto','unitDescription','unitLabel'
    ];
    if (item && Array.isArray(item.units) && item.units.length) {
      item.units.forEach(u=>{
        if (u) {
          if (u.name) candidates.push(String(u.name));
          if (u.number) candidates.push(String(u.number));
          if (u.code) candidates.push(String(u.code));
        }
      });
    }
    fields.forEach(k=>{
      if (item && item[k] !== undefined && item[k] !== null && String(item[k]).trim() !== '') candidates.push(String(item[k]));
      if (baseRow && baseRow[k] !== undefined && baseRow[k] !== null && String(baseRow[k]).trim() !== '') candidates.push(String(baseRow[k]));
    });
    if (!candidates.length) return '';
    const joined = candidates.join(' | ');
    const nums = _extractCandidates(joined);
    return nums.length ? nums.join(' | ') : '';
  }

  // extrai do documentNumber (apenas para adiantamento e com validação de comprimento)
  function _extractFromDocumentNumberWhenAdiantamento(item, baseRow) {
    const docTypeName = _getDocumentTypeName(item, baseRow);
    if (!_isAdiantamentoType(docTypeName)) return '';
    const docNumRaw = (item && item.documentNumber) ? String(item.documentNumber) : (baseRow && baseRow['documentNumber'] ? String(baseRow['documentNumber']) : '');
    if (!docNumRaw || !docNumRaw.trim()) return '';
    const cand = _extractCandidates(docNumRaw);
    const accepted = cand.filter(n => {
      const L = String(n).length;
      return L >= MIN_ACCEPT_LEN && L <= MAX_ACCEPT_LEN;
    });
    return accepted.length ? accepted.join(' | ') : '';
  }

  // === Loop principal (mantendo lógica de receipts) ===
  (data || []).forEach(item => {
    const base = {};
    copyScalarsAndOneLevel(item, '', base, headersSet);

    if (!('installmentId' in base)) base.installmentId = item.installmentId || '';
    if (!('billId' in base)) base.billId = item.billId || '';
    if (!('clientId' in base)) base.clientId = item.clientId || '';
    headersSet.add('installmentId'); headersSet.add('billId'); headersSet.add('clientId');

    // tenta unidade explícita primeiro
    let guessedUnitLabel = _extractFromUnitsFields(item, base);
    // se não houver e for adiantamento, tenta extrair do documentNumber
    if (!guessedUnitLabel) guessedUnitLabel = _extractFromDocumentNumberWhenAdiantamento(item, base);

    headersSet.add('unitNumberExtracted');

    const receipts = Array.isArray(item.receipts) ? item.receipts : [];
    if (receipts.length) {
      receipts.forEach(r => {
        const row = { ...base };
        const rCopy = { ...r };
        delete rCopy.bankMovements;
        copyScalarsAndOneLevel(rCopy, 'receipt', row, headersSet);
        ['receipt.calculationDate','receipt.paymentDate'].forEach(k => { if (row[k]) row[k] = toDateStr_(row[k]); });
        row['unitNumberExtracted'] = guessedUnitLabel || '';
        rows.push(row);
      });
    } else {
      base['unitNumberExtracted'] = guessedUnitLabel || '';
      rows.push(base);
    }
  });

  // === Preenchimento por cliente (fallback) ===
  // mapeia clientKey -> valor mais frequente de unitNumberExtracted
  const freqByClient = {}; // clientKey -> { unit -> count }
  rows.forEach(r => {
    const clientKey = (r['clientId'] !== '' && r['clientId'] !== undefined && r['clientId'] !== null)
      ? String(r['clientId'])
      : (r['clientName'] ? String(r['clientName']).trim().toLowerCase() : '');
    const unit = r['unitNumberExtracted'] ? String(r['unitNumberExtracted']).trim() : '';
    if (!clientKey) return;
    if (!unit) return;
    freqByClient[clientKey] = freqByClient[clientKey] || {};
    freqByClient[clientKey][unit] = (freqByClient[clientKey][unit] || 0) + 1;
  });
  // determina o mais frequente por cliente
  const chosenByClient = {};
  Object.keys(freqByClient).forEach(cli=>{
    const map = freqByClient[cli];
    let best = '', bestCount = 0;
    Object.keys(map).forEach(u=>{
      if (map[u] > bestCount) { best = u; bestCount = map[u]; }
    });
    if (best) chosenByClient[cli] = best;
  });

  // aplica preenchimento: quando unitNumberExtracted vazia e existe chosenByClient para esse cliente
  let fillCount = 0;
  rows.forEach(r=>{
    if (r['unitNumberExtracted'] && String(r['unitNumberExtracted']).trim() !== '') return;
    const clientKey = (r['clientId'] !== '' && r['clientId'] !== undefined && r['clientId'] !== null)
      ? String(r['clientId'])
      : (r['clientName'] ? String(r['clientName']).trim().toLowerCase() : '');
    if (!clientKey) return;
    const chosen = chosenByClient[clientKey];
    if (chosen) {
      r['unitNumberExtracted'] = chosen;
      fillCount++;
    }
  });
  console.log(`ℹ️ Preenchimento automático de Número da Unidade por cliente: ${fillCount} células preenchidas.`);

  const preferred = ['installmentId','billId','clientId','unitNumberExtracted'];
  const headers = preferred.concat(Array.from(headersSet).filter(h => !preferred.includes(h)).sort());
  return { headers, rows };
}

// ========== ESCRITA (Ajuste de grade + cabeçalhos PT-BR) ==========
function escreverNaPlanilhaIncomeFlat_(headersPt, headersRaw, dados) {
  if (typeof Sheets !== 'undefined') {
    escreverComSheetsApi_(headersPt, headersRaw, dados);
  } else {
    console.log('ℹ️ Sheets API não habilitado — usando fallback SpreadsheetApp.');
    escreverComSpreadsheetApp_(headersPt, headersRaw, dados);
  }
}

// ---------- Via Google Sheets Advanced Service ----------
function callWithRetry_(fn, tries, baseMs) {
  let lastErr;
  for (let i = 0; i < tries; i++) {
    try { return fn(); } catch (e) {
      lastErr = e;
      Utilities.sleep(baseMs * Math.pow(2, i) + Math.floor(Math.random()*200));
    }
  }
  throw lastErr;
}

function normalizeName_(s) {
  return String(s || '').trim().toLowerCase();
}

// Retorna { sheetId, titleExato }
function ensureSheet_(spreadsheetId, sheetNameRef) {
  const target = normalizeName_(sheetNameRef);
  const ss = callWithRetry_(() => Sheets.Spreadsheets.get(spreadsheetId), 3, 500);

  const found = (ss.sheets || []).find(sh => {
    const title = sh.properties?.title || '';
    return normalizeName_(title) === target;
  });
  if (found) return { sheetId: found.properties.sheetId, titleExato: found.properties.title };

  // cria com o nome de referência
  const req = { requests: [{ addSheet: { properties: { title: sheetNameRef } } }] };
  const res = callWithRetry_(() => Sheets.Spreadsheets.batchUpdate(req, spreadsheetId), 3, 500);
  const addSheetReply = (res.replies || []).find(r => r.addSheet);
  if (addSheetReply?.addSheet?.properties) {
    return {
      sheetId: addSheetReply.addSheet.properties.sheetId,
      titleExato: addSheetReply.addSheet.properties.title
    };
  }
  throw new Error(`Falha ao garantir a aba "${sheetNameRef}".`);
}

function ensureGridSizeSheetsApi_(spreadsheetId, sheetId, needRows, needCols) {
  const ss = callWithRetry_(() => Sheets.Spreadsheets.get(spreadsheetId), 3, 500);
  const sh = (ss.sheets || []).find(s => s.properties?.sheetId === sheetId);
  const curRows = sh?.properties?.gridProperties?.rowCount || 1000;
  const curCols = sh?.properties?.gridProperties?.columnCount || 26;

  const newRows = Math.max(curRows, needRows);
  const newCols = Math.max(curCols, needCols);
  if (newRows === curRows && newCols === curCols) return;

  const req = {
    requests: [{
      updateSheetProperties: {
        properties: { sheetId, gridProperties: { rowCount: newRows, columnCount: newCols } },
        fields: 'gridProperties.rowCount,gridProperties.columnCount'
      }
    }]
  };
  callWithRetry_(() => Sheets.Spreadsheets.batchUpdate(req, spreadsheetId), 3, 500);
}

function clearSheetValuesSheetsApi_(spreadsheetId, titleExato) {
  callWithRetry_(() => Sheets.Spreadsheets.Values.clear({}, spreadsheetId, titleExato), 3, 500);
}

function escreverComSheetsApi_(headersPt, headersRaw, dados) {
  const spreadsheetId = CONFIG_INCOME.SPREADSHEET_ID;

  // garante aba e usa o TÍTULO EXATO encontrado
  const { sheetId, titleExato } = ensureSheet_(spreadsheetId, CONFIG_INCOME.SHEET_NAME);

  // calcula necessidades
  const totalRowsNeeded = 1 + (dados?.length || 0);
  const totalColsNeeded = headersPt.length;

  // dimensiona (para cima) antes de escrever
  ensureGridSizeSheetsApi_(spreadsheetId, sheetId, totalRowsNeeded + 50, totalColsNeeded);

  // limpa valores
  clearSheetValuesSheetsApi_(spreadsheetId, titleExato);

  if (!dados || !dados.length) {
    Sheets.Spreadsheets.Values.update(
      { values: [['Sem dados']] },
      spreadsheetId,
      `${titleExato}!A1`,
      { valueInputOption: 'RAW' }
    );
  } else {
    // cabeçalho PT-BR
    callWithRetry_(() => Sheets.Spreadsheets.Values.update(
      { values: [headersPt] },
      spreadsheetId,
      `${titleExato}!A1`,
      { valueInputOption: 'RAW' }
    ), 3, 400);

    // blocos
    const COLS = headersRaw.length;
    const CELLS_PER_CHUNK = 120000;
    const ROWS_PER_CHUNK = Math.max(200, Math.floor(CELLS_PER_CHUNK / COLS));

    let startRow = 2;
    for (let i = 0; i < dados.length; i += ROWS_PER_CHUNK) {
      const slice = dados.slice(i, i + ROWS_PER_CHUNK);
      const values = slice.map(obj => headersRaw.map(h => normalizeCell_(obj[h])));

      // garante grade suficiente para este bloco
      ensureGridSizeSheetsApi_(spreadsheetId, sheetId, startRow - 1 + values.length + 50, headersPt.length);

      callWithRetry_(() => Sheets.Spreadsheets.Values.update(
        { values },
        spreadsheetId,
        `${titleExato}!A${startRow}`,
        { valueInputOption: 'RAW' }
      ), 3, 500);

      startRow += values.length;
    }
  }

  // congela header
  callWithRetry_(() => Sheets.Spreadsheets.batchUpdate({
    requests: [
      { updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } }
    ]
  }, spreadsheetId), 3, 600);

  // ====> FORÇA "DATA DE VENCIMENTO" COMO DATA (Sheets API)
  enforceDueDateAsDateSheetsApi_(
    spreadsheetId,
    sheetId,
    titleExato,
    headersPt,
    headersRaw,
    dados || []
  );

  // encolhe para o número exato de linhas/colunas (apaga colunas removidas)
  if (CONFIG_INCOME.SHRINK_SHEET) {
    resizeGridExactSheetsApi_(spreadsheetId, sheetId, Math.max(2, totalRowsNeeded), Math.max(1, totalColsNeeded));
  }
}

// ---------- Fallback via SpreadsheetApp ----------
function escreverComSpreadsheetApp_(headersPt, headersRaw, dados) {
  const ss  = SpreadsheetApp.openById(CONFIG_INCOME.SPRЕADSHЕЕT_ID || CONFIG_INCOME.SPREADSHEET_ID); // tolera typo antigo

  // encontra aba por nome normalizado (ignora caixa)
  const target = normalizeName_(CONFIG_INCOME.SHEET_NAME);
  let sheet = ss.getSheets().find(sh => normalizeName_(sh.getName()) === target);
  if (!sheet) sheet = ss.insertSheet(CONFIG_INCOME.SHEET_NAME);

  sheet.clear();

  const needRows = 1 + (dados?.length || 0);
  const needCols = headersPt.length;

  // dimensiona para cima
  if (sheet.getMaxRows() < needRows + 50) sheet.insertRowsAfter(sheet.getMaxRows(), needRows + 50 - sheet.getMaxRows());
  if (sheet.getMaxColumns() < needCols)  sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols - sheet.getMaxColumns());

  if (!dados || !dados.length) {
    sheet.getRange(1,1).setValue('Sem dados');
  } else {
    // header PT-BR
    sheet.getRange(1, 1, 1, headersPt.length).setValues([headersPt]);

    // blocos
    const COLS = headersRaw.length;
    const CELLS_PER_CHUNK = 60000;
    const ROWS_PER_CHUNK = Math.max(200, Math.floor(CELLS_PER_CHUNK / COLS));

    let startRow = 2;
    for (let i = 0; i < dados.length; i += ROWS_PER_CHUNK) {
      const slice = dados.slice(i, i + ROWS_PER_CHUNK);
      const values = slice.map(obj => headersRaw.map(h => normalizeCell_(obj[h])));
      const range = sheet.getRange(startRow, 1, values.length, headersPt.length);
      range.setValues(values);
      startRow += values.length;
      SpreadsheetApp.flush();
      Utilities.sleep(60);
    }
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headersPt.length);

  // ====> FORÇA "DATA DE VENCIMENTO" COMO DATA (SpreadsheetApp) — ROBUSTA
  const totalRowsWritten = 1 + (dados?.length || 0);
  enforceDueDateAsDateSpreadsheetApp_(sheet, headersPt, totalRowsWritten);

  // encolhe para o número exato de linhas/colunas
  if (CONFIG_INCOME.SHRINK_SHEET) {
    const maxCols = sheet.getMaxColumns();
    const maxRows = sheet.getMaxRows();
    if (maxCols > needCols) sheet.deleteColumns(needCols + 1, maxCols - needCols);
    if (maxRows > needRows + 2) sheet.deleteRows(needRows + 3, maxRows - (needRows + 2)); // deixa um respiro
  }
}

// ========== HELPERS ==========
function toDateStr_(v) {
  if (!v) return '';
  const s = String(v).slice(0,10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function normalizeCell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v); // segurança (não deve ocorrer após flatten)
  return v;
}

// Remove colunas com base em lista de chaves e/ou prefixos (raw keys)
function applyColumnExclusions_(headersRaw, rows, excludeKeys, excludePrefixes) {
  const excludes = new Set(excludeKeys || []);
  const prefixes = (excludePrefixes || []).slice();

  function keepKey(k) {
    if (excludes.has(k)) return false;
    for (const p of prefixes) {
      if (k === p || k.startsWith(p + '.') || k.startsWith(p)) return false;
    }
    return true;
  }

  const filteredHeaders = headersRaw.filter(keepKey);
  const filteredRows = rows.map(r => {
    const o = {};
    filteredHeaders.forEach(h => { o[h] = r[h]; });
    return o;
  });

  return { headers: filteredHeaders, rows: filteredRows };
}

// Redimensiona EXATAMENTE a grade (Sheets API)
function resizeGridExactSheetsApi_(spreadsheetId, sheetId, rows, cols) {
  callWithRetry_(() => Sheets.Spreadsheets.batchUpdate({
    requests: [{
      updateSheetProperties: {
        properties: { sheetId, gridProperties: { rowCount: rows, columnCount: cols } },
        fields: 'gridProperties.rowCount,gridProperties.columnCount'
      }
    }]
  }, spreadsheetId), 3, 500);
}

/**
 * Tradutor de cabeçalhos para PT-BR, gerando nomes únicos e consistentes.
 */
function translateKeyPt_(key) {
  const map = {
    // chaves críticas
    'installmentId': 'Parcela - ID',
    'billId': 'Título - ID',
    'clientId': 'Cliente - ID',
    'clientName': 'Cliente - Nome',

    // nossa coluna derivada de unidade
    'unitNumberExtracted': 'Número da Unidade',

    // empresa/grupo/projeto
    'companyId': 'Empresa - ID',
    'companyName': 'Empresa - Nome',
    'groupCompanyId': 'Grupo - ID',
    'groupCompanyName': 'Grupo - Nome',
    'holdingId': 'Holding ID',
    'holdingName': 'Holding Nome',
    'subsidiaryId': 'Subsidiária ID',
    'subsidiaryName': 'Subsidiária Nome',
    'projectId': 'Projeto ID',
    'projectName': 'Projeto - Nome',
    'businessAreaId': 'Negócio Área ID',
    'businessAreaName': 'Negócio Área Nome',
    'businessTypeId': 'Negócio Tipo ID',
    'businessTypeName': 'Negócio Tipo Nome',

    // documento
    'documentIdentificationId': 'Tipo de Documento - ID',
    'documentIdentificationName': 'Tipo de Documento - Nome',
    'documentNumber': 'Documento - Número',
    'documentForecast': 'Documento - Previsão (S/N)',

    // origem/portador/unidade
    'originId': 'Origem - ID',
    'bearerId': 'Portador - ID',
    'mainUnit': 'Unidade Principal',
    'installmentNumber': 'Parcela - Nº',

    // datas parcela
    'issueDate': 'Data de Emissão',
    'dueDate': 'Data de Vencimento',
    'billDate': 'Data de Competência',
    'installmentBaseDate': 'Parcela - Data Base',
    'interestBaseDate': 'Juros - Data Base',

    // valores parcela
    'originalAmount': 'Valor Original',
    'discountAmount': 'Valor Desconto (Título)',
    'taxAmount': 'Valor Imposto Retido',
    'balanceAmount': 'Saldo',
    'correctedBalanceAmount': 'Saldo Corrigido',
    'embeddedInterestAmount': 'Embutidos Juros Valor',

    // indexação/juros
    'indexerId': 'Indexador - ID',
    'indexerName': 'Indexador - Nome',
    'periodicityType': 'Periodicidade - Tipo',
    'interestType': 'Juros - Tipo',
    'interestRate': 'Juros - Percentual',
    'correctionType': 'Correção - Tipo',

    // situação e condição de pagamento
    'defaulterSituation': 'Situação de Inadimplência',
    'subJudicie': 'Sub Judicie',
    'paymentTerm.id': 'Condição de Pagamento - ID',
    'paymentTerm.description': 'Condição de Pagamento - Descrição',

    // recibo
    'receipt.sequencialNumber': 'Recibo - Sequencial',
    'receipt.operationTypeId': 'Recibo - Tipo de Operação - ID',
    'receipt.operationTypeName': 'Recibo - Tipo de Operação - Nome',
    'receipt.grossAmount': 'Recibo - Valor Bruto',
    'receipt.monetaryCorrectionAmount': 'Recibo - Correção Monetária',
    'receipt.interestAmount': 'Recibo - Juros',
    'receipt.fineAmount': 'Recibo - Multa',
    'receipt.discountAmount': 'Recibo - Desconto',
    'receipt.taxAmount': 'Recibo - Imposto Valor',
    'receipt.netAmount': 'Recibo - Valor Líquido',
    'receipt.additionAmount': 'Recibo - Acréscimo',
    'receipt.insuranceAmount': 'Recibo - Seguro',
    'receipt.dueAdmAmount': 'Recibo - Taxa Adm',
    'receipt.embeddedInterestAmount': 'Recibo - Juros Embutidos',
    'receipt.proRata': 'Recibo - Pro Rata',
    'receipt.calculationDate': 'Recibo - Data de Cálculo',
    'receipt.paymentDate': 'Recibo - Data de Pagamento',
    'receipt.accountCompanyId': 'Recibo - Conta (Empresa) - ID',
    'receipt.accountNumber': 'Recibo - Conta - Número',
    'receipt.accountType': 'Recibo - Conta - Tipo',
    'receipt.indexerId': 'Recibo - Indexador - ID'
  };

  if (map[key]) return map[key];

  // heurística genérica
  if (key.startsWith('receipt.')) {
    return 'Recibo - ' + humanizePt_(key.slice('receipt.'.length));
  }
  if (key.startsWith('paymentTerm.')) {
    return 'Condição de Pagamento - ' + humanizePt_(key.slice('paymentTerm.'.length));
  }
  if (key.startsWith('documentIdentification.')) {
    return 'Tipo de Documento - ' + humanizePt_(key.slice('documentIdentification.'.length));
  }
  return humanizePt_(key);
}

// Converte camelCase/dot em rótulos PT-BR simples e estáveis
function humanizePt_(raw) {
  const parts = String(raw).split('.');
  const last = parts.pop();
  const head = parts.map(p => capitalizePt_(splitCamel_(p).join(' '))).join(' - ');
  const tail = splitCamel_(last).map(trPt_).map(capitalizePt_).join(' ');
  return head ? `${capitalizePt_(head)} - ${tail}` : tail;
}
function splitCamel_(s) {
  return String(s)
    .replace(/[_\-]+/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .toLowerCase()
    .trim()
    .split(/\s+/);
}
function trPt_(w) {
  const dict = {
    id: 'ID', name: 'Nome', number: 'Número', date: 'Data', amount: 'Valor',
    original: 'Original', discount: 'Desconto', tax: 'Imposto', balance: 'Saldo',
    corrected: 'Corrigido', embedded: 'Embutidos', interest: 'Juros', rate: 'Percentual',
    type: 'Tipo', issue: 'Emissão', due: 'Vencimento', bill: 'Competência', base: 'Base',
    document: 'Documento', forecast: 'Previsão', indexer: 'Indexador', payment: 'Pagamento',
    bearer: 'Portador', situation: 'Situação', defaulter: 'Inadimplência', periodicity: 'Periodicidade',
    correction: 'Correção', company: 'Empresa', client: 'Cliente', group: 'Grupo', holding: 'Holding',
    project: 'Projeto', area: 'Área', business: 'Negócio', unit: 'Unidade', main: 'Principal',
    subsidiary: 'Subsidiária'
  };
  return dict[w] || w;
}
function capitalizePt_(s) {
  if (s === 'ID') return s;
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/* =========================
   === HELPERS ADICIONAIS ===
   ========================= */

// Converte índice (1-based) em letra de coluna A1 ("A", "B", ..., "AA")
function indexToA1Col_(idx1) {
  let n = idx1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Força a coluna "Data de Vencimento" a ser interpretada como DATA (via Sheets API)
function enforceDueDateAsDateSheetsApi_(spreadsheetId, sheetId, titleExato, headersPt, headersRaw, dados) {
  const ptIdx = headersPt.indexOf('Data de Vencimento'); // 0-based
  const rawIdx = headersRaw.indexOf('dueDate');          // 0-based
  if (ptIdx < 0 || rawIdx < 0 || !dados || !dados.length) return;

  const col1Based = ptIdx + 1;
  const colA1 = indexToA1Col_(col1Based);
  const startRow = 2;
  const endRow = 1 + dados.length;
  const rangeA1 = `${titleExato}!${colA1}${startRow}:${colA1}${endRow}`;

  // prepara valores ISO (yyyy-MM-dd) para USER_ENTERED
  const colValues = dados.map(obj => {
    const s = toDateStr_(obj['dueDate']);
    return [s || ''];
  });

  // Reescreve a coluna como USER_ENTERED (Sheets interpreta data)
  callWithRetry_(() => Sheets.Spreadsheets.Values.update(
    { values: colValues },
    spreadsheetId,
    rangeA1,
    { valueInputOption: 'USER_ENTERED' }
  ), 3, 400);

  // Aplica formato de data (ajuste o pattern se quiser 'dd/MM/yyyy')
  callWithRetry_(() => Sheets.Spreadsheets.batchUpdate({
    requests: [{
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: startRow - 1,
          endRowIndex: endRow,
          startColumnIndex: col1Based - 1,
          endColumnIndex: col1Based
        },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }
          }
        },
        fields: 'userEnteredFormat.numberFormat'
      }
    }]
  }, spreadsheetId), 3, 500);
}

// Força a coluna "Data de Vencimento" a ser DATA (via SpreadsheetApp) — versão robusta
function enforceDueDateAsDateSpreadsheetApp_(sheet, headersPt, totalRows) {
  const colIdx0 = headersPt.indexOf('Data de Vencimento'); // 0-based
  if (colIdx0 < 0 || totalRows <= 1) return;

  const col = colIdx0 + 1; // 1-based
  const rows = totalRows - 1;
  const range = sheet.getRange(2, col, rows, 1);
  const vals = range.getValues(); // [[v], [v], ...]

  function tryParseDate(v) {
    if (v === '' || v === null || v === undefined) return null;

    // 1) Já é Date?
    if (v instanceof Date && !isNaN(v.getTime())) return v;

    // 2) Número (serial Sheets)
    if (typeof v === 'number') {
      const epoch = new Date(1899, 11, 30); // base do serial do Sheets
      const out = new Date(epoch.getTime() + v * 24 * 3600 * 1000);
      return isNaN(out.getTime()) ? null : out;
    }

    // 3) String — normaliza e tenta vários formatos
    const s = String(v).trim();

    // ISO puro yyyy-mm-dd
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      const [_, Y, M, D] = m.map(Number);
      const d = new Date(Y, M - 1, D);
      return isNaN(d.getTime()) ? null : d;
    }

    // ISO com tempo: yyyy-mm-ddTHH:mm(:ss.sss)?(Z|±hh:mm)?
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T ]/);
    if (m) {
      const [_, Y, M, D] = m.map(Number);
      const d = new Date(Y, M - 1, D);
      return isNaN(d.getTime()) ? null : d;
    }

    // BR: dd/MM/yyyy
    m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
      const D = Number(m[1]), M = Number(m[2]), Y = Number(m[3]);
      const d = new Date(Y, M - 1, D);
      return isNaN(d.getTime()) ? null : d;
    }

    // US: MM/dd/yyyy
    m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
      const M = Number(m[1]), D = Number(m[2]), Y = Number(m[3]);
      const d = new Date(Y, M - 1, D);
      return isNaN(d.getTime()) ? null : d;
    }

    return null;
  }

  const out = vals.map(([v]) => {
    const parsed = tryParseDate(v);
    // Se não der pra interpretar, preserva o original (não apaga!)
    return [parsed ?? v];
  });

  range.setValues(out);
  // Formato; ajuste se quiser brasileiro: 'dd/MM/yyyy'
  range.setNumberFormat('yyyy-mm-dd');
}
