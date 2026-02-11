/**
 * APPS SCRIPT — UNIDADES (endpoint /units) - VERSÃO AJUSTADA
 * Aba: "Unidades"
 *
 * Configure Script Properties:
 *  - SIENGE_SUBDOMAIN
 *  - SIENGE_USER
 *  - SIENGE_PASS
 *  - SPREADSHEET_ID
 *
 * Ajuste ENTERPRISE_ID abaixo se precisar.
 */

// ======== CONFIG ========
const CONFIG_UNITS = {
  SIENGE_SUBDOMAIN: PropertiesService.getScriptProperties().getProperty('SIENGE_SUBDOMAIN') || 'm21',
  SIENGE_USER:      PropertiesService.getScriptProperties().getProperty('SIENGE_USER'),
  SIENGE_PASS:      PropertiesService.getScriptProperties().getProperty('SIENGE_PASS'),
  ENTERPRISE_ID:    78,                         // centro de custo
  SPREADSHEET_ID:   PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  SHEET_NAME:       'Unidades',
  PAGE_LIMIT:       200,                        // conforme swagger
  ADDITIONAL_DATA:  'ALL',                      // ou 'NONE' ou null
  // forçar todos os códigos de estoque comercial para tentar puxar "tudo"
  COMMERCIAL_STOCK_IN: ['C','D','R','E','M','P','V','L','T','G','O']
};

// ======== MAPEAMENTO commercialStock ========
const COMMERCIAL_STOCK_LABELS = {
  'C': 'Reservada',
  'D': 'Disponível',
  'R': 'Reserva técnica',
  'E': 'Permuta',
  'M': 'Mútuo',
  'P': 'Proposta',
  'V': 'Vendida',
  'L': 'Locado',
  'T': 'Transferido',
  'G': 'Vendido/terceiros',
  'O': 'Vendida em pré-contrato'
};

// ======== HELPERS ========
function _safeNumber(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return isFinite(v) ? v : '';
  let s = String(v).trim();
  if (!s) return '';
  s = s.replace(/[^\d\-,.]/g,'');
  if (!s) return '';
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    s = s.replace(/\./g,'').replace(',', '.');
  } else if (s.indexOf(',') > -1) {
    s = s.replace(',', '.');
  }
  const n = Number(s);
  return isNaN(n) ? '' : n;
}
function _safeString(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}
function _joinChildUnits(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return '';
  return arr.map(u => {
    const id = u.id ?? u.unitId ?? '';
    const name = u.name ?? u.unitName ?? u.label ?? '';
    return [id, name].filter(Boolean).join(' - ');
  }).join(' | ');
}
function _joinGroupings(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return '';
  return arr.map(g => {
    const gd = g.groupingDescription ?? '';
    const vg = g.valueGroupingDescription ?? '';
    return [gd, vg].filter(Boolean).join(': ');
  }).join(' | ');
}
function _joinSpecialValues(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return '';
  return arr.map(sv => {
    const t = sv.tablePricesID ?? sv.tablePricesId ?? '';
    const iq = sv.indexedQuantity ?? '';
    return `tableId:${t}${ iq !== '' ? ' q:' + iq : '' }`;
  }).join(' | ');
}
function _firstNonEmpty(...args) {
  for (const a of args) if (a !== undefined && a !== null && a !== '') return a;
  return '';
}

// converte dígito '1'..'4' para romano I..IV
function _digitToRoman(d) {
  const map = { '1': 'I', '2': 'II', '3': 'III', '4': 'IV' };
  return map[String(d)] || '';
}

// ======== FUNÇÃO PRINCIPAL ========
function atualizarUnidades() {
  try {
    console.log('🚀 Iniciando Unidades...');
    const dados = fetchAllUnits();
    const linhas = processUnitsToRows(dados);
    escreverNaPlanilhaUnidades(linhas);
    console.log(`✅ Concluído: ${linhas.length} linha(s) escritas em "${CONFIG_UNITS.SHEET_NAME}"`);
  } catch (e) {
    console.error('❌ Erro em atualizarUnidades:', e);
  }
}

// ======== FETCH (paginação offset/limit) ========
function fetchAllUnits() {
  const base = `https://api.sienge.com.br/${CONFIG_UNITS.SIENGE_SUBDOMAIN}/public/api/v1/units`;
  const limit = CONFIG_UNITS.PAGE_LIMIT || 200;
  let offset = 0;
  const all = [];

  // montar lista comercialStockIn (string csv) se definida
  const commercialStockParam = Array.isArray(CONFIG_UNITS.COMMERCIAL_STOCK_IN) && CONFIG_UNITS.COMMERCIAL_STOCK_IN.length
    ? CONFIG_UNITS.COMMERCIAL_STOCK_IN.join(',')
    : null;

  while (true) {
    const params = [];
    params.push('limit=' + encodeURIComponent(limit));
    params.push('offset=' + encodeURIComponent(offset));
    if (CONFIG_UNITS.ENTERPRISE_ID) params.push('enterpriseId=' + encodeURIComponent(CONFIG_UNITS.ENTERPRISE_ID));
    if (CONFIG_UNITS.ADDITIONAL_DATA) params.push('additionalData=' + encodeURIComponent(CONFIG_UNITS.ADDITIONAL_DATA));
    if (commercialStockParam) params.push('commercialStockIn=' + encodeURIComponent(commercialStockParam)); // força todos os estoques
    const url = base + '?' + params.join('&');

    const options = {
      method: 'GET',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(`${CONFIG_UNITS.SIENGE_USER}:${CONFIG_UNITS.SIENGE_PASS}`),
        Accept: 'application/json'
      },
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      timeout: 60000
    };

    console.log(`📡 GET (offset ${offset}): ${url}`);
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    const txt = resp.getContentText();
    if (code !== 200) {
      throw new Error(`HTTP ${code} :: ${txt}`);
    }

    let json;
    try { json = JSON.parse(txt); } catch (err) {
      // tentar limpar caracteres estranhos e parsear
      try { json = JSON.parse(txt.replace(/[\u0000-\u001F]+/g,'')); }
      catch (err2) {
        console.error('❌ Falha ao parsear JSON (trecho):', txt.slice(0,1000));
        throw new Error('Falha ao parsear JSON da API. Veja logs.');
      }
    }

    // extrair array de resultados robusto
    let pageResults = [];
    if (Array.isArray(json.results)) {
      pageResults = json.results;
    } else if (Array.isArray(json)) {
      pageResults = json;
    } else if (Array.isArray(json.data)) {
      pageResults = json.data;
    } else if (Array.isArray(json.content)) {
      pageResults = json.content;
    } else {
      // procurar primeiro array no objeto
      for (const v of Object.values(json)) {
        if (Array.isArray(v)) { pageResults = v; break; }
      }
    }

    const count = json.resultSetMetadata?.count ?? (Array.isArray(pageResults) ? pageResults.length : 0);
    console.log(`📦 Página offset ${offset} retornou ${pageResults.length} unidade(s) — total dispon.: ${count}`);
    all.push(...pageResults);

    // controle de paginação
    if (json.resultSetMetadata && json.resultSetMetadata.count !== undefined && json.resultSetMetadata.offset !== undefined && json.resultSetMetadata.limit !== undefined) {
      const nextOffset = json.resultSetMetadata.offset + json.resultSetMetadata.limit;
      if (nextOffset >= json.resultSetMetadata.count) break;
      offset = nextOffset;
    } else {
      // fallback: se pageResults menor que limit -> fim
      if (!pageResults || pageResults.length < limit) break;
      offset += limit;
    }
  }

  console.log(`📦 Total coletado: ${all.length} unidade(s)`);
  return all;
}

// ======== TRANSFORMAR EM LINHAS ========
function processUnitsToRows(units) {
  console.log(`📊 Processando ${units.length} unidade(s)...`);
  const rows = units.map(u => {
    // tentar vários campos possíveis para cada informação (fallbacks)
    const commercialCode = _firstNonEmpty(u.commercialStock, u.commercialStockCode, u.stockCode, u.stock) || '';
    const unitId = _firstNonEmpty(u.id, u.unitId, u.code) || '';
    const enterpriseId = _firstNonEmpty(u.enterpriseId, u.enterprise?.id) || '';
    const enterpriseName = _firstNonEmpty(u.enterpriseName, u.enterprise?.name, u.enterpriseName) || '';
    const contractId = _firstNonEmpty(u.contractId, u.contract?.id) || '';
    const name = _firstNonEmpty(u.name, u.unitName, u.label, u.identification) || '';
    const propertyType = _firstNonEmpty(u.propertyType, u.type, u.property) || '';

    // AGRUPAMENTOS: capturar tanto texto final quanto array original para extrair módulo
    const rawGroupingsArray = _firstNonEmpty(u.groupings, u.groupingsList, []);
    const agrupamentosText = _joinGroupings(rawGroupingsArray);

    // tentativa 1: procurar no array de objetos (groupingDescription/valueGroupingDescription)
    let moduloRoman = '';
    if (Array.isArray(rawGroupingsArray) && rawGroupingsArray.length) {
      for (const g of rawGroupingsArray) {
        const cand = [g.groupingDescription, g.valueGroupingDescription, g.valueGrouping, g.value, g.description].filter(Boolean).join(' ');
        if (cand) {
          // busca por palavra MÓDULO e dígito 1..4 (aceita MODULO sem acento)
          const m = cand.match(/M[ÓO]DULO[^0-9]*([1-4])/i) || cand.match(/MODULO[^0-9]*([1-4])/i);
          if (m && m[1]) { moduloRoman = _digitToRoman(m[1]); break; }
        }
      }
    }

    // tentativa 2: buscar no texto concatenado (fallback)
    if (!moduloRoman && agrupamentosText) {
      const m2 = agrupamentosText.match(/M[ÓO]DULO[^0-9]*([1-4])/i) || agrupamentosText.match(/MODULO[^0-9]*([1-4])/i);
      if (m2 && m2[1]) moduloRoman = _digitToRoman(m2[1]);
    }

    return {
      'ID Unidade': unitId,
      'ID Empreendimento (CC)': enterpriseId,
      'ID Contrato': contractId,
      'Nome / Identificador': _safeString(name),
      'Tipo Imóvel': _safeString(propertyType),
      'Estoque Comercial (Código)': _safeString(commercialCode),
      'Estoque Comercial (Descrição)': COMMERCIAL_STOCK_LABELS[commercialCode] || '',
      'Matrícula / Inscrição': _safeString(_firstNonEmpty(u.legalRegistrationNumber, u.registryNumber, u.registration)),
      'Número do Contrato': _safeString(_firstNonEmpty(u.contractNumber, u.number)),
      'Área Privativa': _safeNumber(_firstNonEmpty(u.privateArea, u.usablePrivateArea)),
      'Agrupamentos': agrupamentosText,
      'Módulo Romano': moduloRoman,
      'Links': Array.isArray(u.links) ? u.links.map(l => `${l.rel}:${l.href}`).join(' | ') : ''
    };
  });

  // ordenar por ID Unidade (numérico, menor -> maior). converte vazios para Infinity pra irem pro final
  rows.sort((a,b) => {
    const na = a['ID Unidade'] === '' ? Infinity : Number(String(a['ID Unidade']).replace(/[^\d-]/g,''));
    const nb = b['ID Unidade'] === '' ? Infinity : Number(String(b['ID Unidade']).replace(/[^\d-]/g,''));
    if (isNaN(na) && isNaN(nb)) return 0;
    if (isNaN(na)) return 1;
    if (isNaN(nb)) return -1;
    return na - nb;
  });

  console.log(`✅ Linhas processadas: ${rows.length}`);
  return rows;
}

// ======== ESCREVER NA PLANILHA (em lotes) ========
function escreverNaPlanilhaUnidades(dados) {
  const ss = CONFIG_UNITS.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG_UNITS.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_UNITS.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_UNITS.SHEET_NAME);
  } else {
    // limpa só o conteúdo mas mantém a aba (evita erro "já existe página")
    sheet.clear();
  }

  if (!dados.length) {
    sheet.getRange(1,1).setValue('Sem dados');
    console.log('⚠️ Nenhum dado para escrever.');
    return;
  }

  // cabeçalho fixo para evitar colunas vazias aleatórias
  const headers = [
    'ID Unidade','ID Empreendimento (CC)','ID Contrato','Nome / Identificador','Tipo Imóvel',
    'Estoque Comercial (Código)','Estoque Comercial (Descrição)','Matrícula / Inscrição','Número do Contrato',
    'Área Privativa','Agrupamentos','Módulo Romano','Links'
  ];

  const matrix = [headers];
  dados.forEach(o => matrix.push(headers.map(h => o[h] === undefined || o[h] === null ? '' : o[h])));

  // escrita em chunks para evitar timeout / limits
  const chunkSize = 500;
  let rowStart = 1;
  for (let i = 0; i < matrix.length; i += chunkSize) {
    const chunk = matrix.slice(i, i + chunkSize);
    const range = sheet.getRange(rowStart, 1, chunk.length, headers.length);
    range.setValues(chunk);
    rowStart += chunk.length;
  }

  // cabeçalho style
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);

  // Formatos numéricos
  const numericCols = ['Área Privativa','Área Útil','Área Comum','Área Terreno','Fração Ideal','Fração Ideal (m2)','Fração VGV','Quantidade Indexada','Valor Terreno','IPTU (R$)','Valor Adimplência Premiada'];
  const lastRow = sheet.getLastRow();
  const idxMap = headers.reduce((acc,h,i)=>{ acc[h]=i+1; return acc; }, {});
  numericCols.forEach(h=>{
    const c = idxMap[h];
    if (c && lastRow > 1) {
      try { sheet.getRange(2, c, lastRow-1, 1).setNumberFormat('#,##0.00'); } catch(e){}
    }
  });

  // Forçar texto em ids/numeros que não devem virar número formatado
  ['ID Unidade','ID Empreendimento (CC)','ID Contrato','Número do Contrato'].forEach(h=>{
    const c = idxMap[h];
    if (c && lastRow > 1) {
      try { sheet.getRange(2, c, lastRow-1, 1).setNumberFormat('@'); } catch(e){}
    }
  });

  try { sheet.autoResizeColumns(1, headers.length); } catch(e){}
  console.log(`📝 ${dados.length} linhas escritas na aba "${CONFIG_UNITS.SHEET_NAME}"`);
}
