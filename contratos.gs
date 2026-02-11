/**
 * SIENGE · CONTRATOS (compatível com resultSetMetadata/results)
 * Endpoint: /public/api/v1/sales-contracts
 * Aba de destino: "Contratos"
 *
 * Configure Script Properties: SIENGE_SUBDOMAIN, SIENGE_USER, SIENGE_PASS, SPREADSHEET_ID
 */

// ========== CONFIG ==========
const CONFIG_CONTRATOS = {
  SIENGE_SUBDOMAIN: PropertiesService.getScriptProperties().getProperty('SIENGE_SUBDOMAIN') || 'm21',
  SIENGE_USER:      PropertiesService.getScriptProperties().getProperty('SIENGE_USER'),
  SIENGE_PASS:      PropertiesService.getScriptProperties().getProperty('SIENGE_PASS'),
  COMPANY_ID: 2007,        // companyId a executar
  ENTERPRISE_ID: 78,    // centro de custo / empreendimento (enterpriseId)
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || null,
  SHEET_NAME: 'Contratos',
  PAGE_LIMIT: 2000,
  MAX_LOOPS: 200
};

// ========== HELPERS ==========
function _toNumber(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return isNaN(v) ? '' : v;
  let s = String(v).trim();
  if (!s) return '';
  s = s.replace(/[^\d\-,.]/g, '');
  if (s === '') return '';
  if (s.indexOf('.') > -1 && s.indexOf(',') > -1) { s = s.replace(/\./g, '').replace(',', '.'); }
  else if (s.indexOf(',') > -1 && s.indexOf('.') === -1) { s = s.replace(',', '.'); }
  const n = Number(s);
  return isNaN(n) ? '' : n;
}
function _firstNonEmpty(...args) {
  for (const a of args) if (a !== undefined && a !== null && a !== '') return a;
  return '';
}

// ----------------- Novo helper: extrair só o número da unidade a partir do "Unidade (Nome)" --------------
/**
 * Recebe uma label de unidade (ex.: "APTO 2202", "PENTHOUSE 2601", "APTO1301", "2102", "CTP26")
 * e retorna a última sequência de dígitos encontrada (preferência), como string.
 * Se não encontrar dígitos, retorna ''.
 */
function _extractUnitNumberFromLabel(label) {
  if (label === undefined || label === null) return '';
  const s = String(label).trim();
  if (!s) return '';

  // busca todas sequências de dígitos
  const all = s.match(/\d+/g);
  if (all && all.length) {
    // retorna a última (ex.: "BLOCO A APTO 2202" -> "2202"; "CTP26" -> "26")
    return all[all.length - 1];
  }

  // fallback: tentar padrões tipo "APTO(\d+)" etc (embora regex acima já cubra)
  const m = s.match(/(?:APTO|APT|PENTHOUSE|PENT|CTP|CT|AD|AP)\D*?(\d{1,6})/i);
  if (m && m[1]) return m[1];

  return '';
}
// -----------------------------------------------------------------------------------------------

// ========== FETCH PÁGINA ==========
function fetchContractsPage_(params, offset) {
  const base = `https://api.sienge.com.br/${CONFIG_CONTRATOS.SIENGE_SUBDOMAIN}/public/api/v1/sales-contracts`;
  const p = Object.assign({}, params);
  p.limit = CONFIG_CONTRATOS.PAGE_LIMIT;
  p.offset = offset || 0;

  const query = Object.keys(p)
    .map(k => {
      const v = p[k];
      if (v === undefined || v === null || v === '') return '';
      return Array.isArray(v) ? v.map(x => `${encodeURIComponent(k)}=${encodeURIComponent(x)}`).join('&') : `${encodeURIComponent(k)}=${encodeURIComponent(v)}`;
    })
    .filter(Boolean)
    .join('&');

  const url = base + (query ? `?${query}` : '');
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(`${CONFIG_CONTRATOS.SIENGE_USER}:${CONFIG_CONTRATOS.SIENGE_PASS}`),
      Accept: 'application/json'
    },
    muteHttpExceptions: true,
    timeout: 60000
  };

  console.log(`📡 GET (offset ${p.offset}): ${url}`);
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code >= 400) throw new Error(`HTTP ${code} :: ${txt}`);

  let json;
  try { json = JSON.parse(txt); } catch (e) {
    // tenta reconhecer se API retornou array puro
    try { json = JSON.parse(txt.replace(/[\u0000-\u001F]+/g,'')); } catch (e2) {
      console.error('❌ Falha ao parsear JSON (trecho):', txt.slice(0, 1000));
      throw new Error('Falha ao parsear JSON da API. Veja logs.');
    }
  }

  // suporte a diferentes wrappers
  let items = [];
  if (Array.isArray(json.results)) items = json.results;
  else if (Array.isArray(json.content)) items = json.content;
  else if (Array.isArray(json.data)) items = json.data;
  else if (Array.isArray(json.items)) items = json.items;
  else if (Array.isArray(json)) items = json;
  else {
    // tenta localizar primeiro array dentro do objeto
    for (const v of Object.values(json)) {
      if (Array.isArray(v)) { items = v; break; }
    }
  }

  // pegar meta simples se existir
  const meta = json.resultSetMetadata || json.meta || null;
  return { items, meta };
}

// ========== COLETAR (paginado) ==========
function fetchAllContracts_() {
  const params = { companyId: CONFIG_CONTRATOS.COMPANY_ID, enterpriseId: CONFIG_CONTRATOS.ENTERPRISE_ID };
  let offset = 0;
  const limit = CONFIG_CONTRATOS.PAGE_LIMIT;
  const all = [];
  let loop = 0;

  while (loop < CONFIG_CONTRATOS.MAX_LOOPS) {
    loop++;
    const { items, meta } = fetchContractsPage_(params, offset);
    if (!items || items.length === 0) {
      console.log(`📦 Página offset ${offset} retornou 0 contrato(s)`);
      break;
    }
    console.log(`📦 Página offset ${offset} retornou ${items.length} contrato(s)`);
    all.push(...items);
    // se meta disponível e indica total/offset/limit, usa para saber fim
    if (meta && (meta.count !== undefined) && (meta.offset !== undefined) && (meta.limit !== undefined)) {
      const nextOffset = meta.offset + meta.limit;
      if (nextOffset >= meta.count) break;
      offset = nextOffset;
    } else {
      // fallback: se retornou menos que limit, fim
      if (items.length < limit) break;
      offset += limit;
    }
  }

  console.log(`✅ Total coletado: ${all.length} contrato(s)`);

  // ===== FILTRO: manter apenas Situação = "Emitido" (case-insensitive) =====
  const before = all.length;
  const onlyEmitidos = all.filter(c => {
    const sit = String(_firstNonEmpty(c.situation, c.status, '')).trim();
    return /^emitido$/i.test(sit);
  });
  console.log(`🔎 Filtragem situação: antes=${before} / depois_emitido=${onlyEmitidos.length}`);

  if (onlyEmitidos.length > 0) {
    const sample = onlyEmitidos[0];
    const sampleKeys = {
      id: sample.id,
      companyId: sample.companyId,
      enterpriseId: sample.enterpriseId,
      number: sample.number,
      value: sample.value,
      salesContractCustomers: Array.isArray(sample.salesContractCustomers) ? sample.salesContractCustomers.length : 'n/a',
      salesContractUnits: Array.isArray(sample.salesContractUnits) ? sample.salesContractUnits.length : 'n/a',
      paymentConditions: Array.isArray(sample.paymentConditions) ? sample.paymentConditions.length : 'n/a',
      receivableBillId: sample.receivableBillId
    };
    console.log('🔎 Amostra de emitidos (campos-chave):', JSON.stringify(sampleKeys));
  }

  return onlyEmitidos;
}

// ========== PROCESSAR ==========
function processContractsToRows_(items) {
  const rows = [];
  items.forEach(c => {
    // filtros de segurança adicionais (garantia)
    if (CONFIG_CONTRATOS.COMPANY_ID && String(c.companyId) !== String(CONFIG_CONTRATOS.COMPANY_ID)) return;
    if (CONFIG_CONTRATOS.ENTERPRISE_ID && String(c.enterpriseId) !== String(CONFIG_CONTRATOS.ENTERPRISE_ID)) return;

    // cliente: usa salesContractCustomers onde main=true, senão first
    let clientIdRaw = '', clientName = '';
    if (Array.isArray(c.salesContractCustomers) && c.salesContractCustomers.length) {
      const main = c.salesContractCustomers.find(x => x.main === true) || c.salesContractCustomers[0];
      clientIdRaw = _firstNonEmpty(main.id, main.customerId, main.clientId, main.partyId) || '';
      clientName = _firstNonEmpty(main.name, main.clientName, main.customerName) || '';
    } else {
      // fallbacks
      clientIdRaw = _firstNonEmpty(c.clientId, c.customerId, c.customer?.id) || '';
      clientName = _firstNonEmpty(c.clientName, c.customer?.name, c.customerName) || '';
    }

    // coerção para número inteiro (ou '' se não conversível)
    const clientIdNumRaw = _toNumber(clientIdRaw);
    const clientIdFinal = (clientIdNumRaw === '') ? '' : Math.trunc(clientIdNumRaw);

    // unidade: usa salesContractUnits[0]
    let unitIdRaw = '', unitName = '';
    if (Array.isArray(c.salesContractUnits) && c.salesContractUnits.length) {
      const u = c.salesContractUnits[0];
      unitIdRaw = _firstNonEmpty(u.id, u.unitId, u.code, u.number) || '';
      unitName = _firstNonEmpty(u.name, u.unitName, u.description) || unitIdRaw || '';
    } else {
      unitIdRaw = _firstNonEmpty(c.unitId, c.unit?.id, c.apartmentId) || '';
      unitName = _firstNonEmpty(c.unitName, c.unit?.name) || '';
    }

    // --- NOVO: extrair somente o número da coluna "Unidade (Nome)"
    const unidadeNumeroExtraida = _extractUnitNumberFromLabel(unitName); // string com dígitos ou ''
    const unidadeNumeroFinal = (unidadeNumeroExtraida === '') ? '' : ( _toNumber(unidadeNumeroExtraida) === '' ? String(unidadeNumeroExtraida) : Math.trunc(_toNumber(unidadeNumeroExtraida)) );

    // valores:
    const valorTotal = _toNumber(_firstNonEmpty(c.value, c.totalSellingValue, c.totalValue, c.contractValue, c.totalValue));
    let valorSaldo = '';
    let valorPago = '';
    if (Array.isArray(c.paymentConditions) && c.paymentConditions.length) {
      let sumOutstanding = 0;
      let sumPaid = 0;
      let anyOutstanding = false;
      let anyPaid = false;
      c.paymentConditions.forEach(pc => {
        const out = _toNumber(_firstNonEmpty(pc.outstandingBalance, pc.outstanding, pc.openingBalance, pc.outstanding_balance, pc.outstandingBalance));
        const paid = _toNumber(_firstNonEmpty(pc.amountPaid, pc.amount_paid, pc.paidValue, pc.paid_amount, pc.paid));
        if (out !== '') { sumOutstanding += Number(out); anyOutstanding = true; }
        if (paid !== '') { sumPaid += Number(paid); anyPaid = true; }
      });
      if (anyOutstanding) valorSaldo = sumOutstanding;
      if (anyPaid) valorPago = sumPaid;
    }
    if (valorSaldo === '' ) valorSaldo = _toNumber(_firstNonEmpty(c.balanceValue, c.openBalance, c.balance, c.outstandingBalance, c.remainingBalance, c.openingBalance));
    if (valorPago === '' ) valorPago = _toNumber(_firstNonEmpty(c.paidValue, c.amountPaid, c.amount_paid, c.paid_amount));

    // ids de títulos
    const receivableBillId = _firstNonEmpty(c.receivableBillId, c.receivableId, c.receivable_bill_id, c.receivable_bill);
    const cancellationPayableBillId = _firstNonEmpty(c.cancellationPayableBillId, c.cancellation_payable_bill_id, c.cancellationPayableBillId, c.cancellationPayableBillId);

    // montar linha (adicionada a nova coluna 'Unidade (Número)')
    rows.push({
      'ID Empresa': c.companyId ?? '',
      'Empresa': c.companyName ?? '',
      'ID Empreendimento (CC)': c.enterpriseId ?? '',
      'Empreendimento (Nome)': c.enterpriseName ?? '',
      'ID Contrato': c.id ?? c.contractId ?? '',
      'Número do Contrato': c.number ?? c.code ?? '',
      'Data do Contrato': c.contractDate ?? c.date ?? '',
      'Situação': c.situation ?? c.status ?? '',
      // ID Cliente vem como número inteiro ou '' (agora)
      'ID Cliente': clientIdFinal === '' ? '' : clientIdFinal,
      'Nome Cliente': clientName ? String(clientName) : '',
      'Unidade (ID)': unitIdRaw ? ( _toNumber(unitIdRaw) === '' ? String(unitIdRaw) : Math.trunc(_toNumber(unitIdRaw)) ) : '',
      'Unidade (Nome)': unitName ? String(unitName) : '',
      'Unidade (Número)': unidadeNumeroFinal === '' ? '' : unidadeNumeroFinal,
      'Valor Total do Contrato (R$)': (valorTotal === '' ? '' : Number(valorTotal)),
      'Valor Saldo Aberto (R$)': (valorSaldo === '' ? '' : Number(valorSaldo)),
      'Valor Pago (R$)': (valorPago === '' ? '' : Number(valorPago)),
      'ID Título (Recebível)': receivableBillId ? String(receivableBillId) : '',
      'ID Título Cancelamento (a Pagar)': cancellationPayableBillId ? String(cancellationPayableBillId) : ''
    });
  });

  // ordenar por ID Cliente numericamente (menor -> maior). IDs vazios ficam ao final.
  rows.sort((a, b) => {
    const A = (a['ID Cliente'] === '' || a['ID Cliente'] === null || a['ID Cliente'] === undefined) ? Infinity : Number(a['ID Cliente']);
    const B = (b['ID Cliente'] === '' || b['ID Cliente'] === null || b['ID Cliente'] === undefined) ? Infinity : Number(b['ID Cliente']);
    return A - B;
  });

  return rows;
}

// ========== ESCREVER NA PLANILHA ==========
function escreverContratosNaPlanilha_(rows) {
  const ss = CONFIG_CONTRATOS.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG_CONTRATOS.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_CONTRATOS.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_CONTRATOS.SHEET_NAME); else sheet.clear();

  if (!rows || rows.length === 0) {
    sheet.getRange(1,1).setValue('Sem dados');
    console.log('⚠️ Nenhum dado para escrever.');
    return;
  }

  // NOTE: adicionei a nova coluna 'Unidade (Número)' entre 'Unidade (Nome)' e valores.
  const headers = [
    'ID Empresa','Empresa','ID Empreendimento (CC)','Empreendimento (Nome)','ID Contrato',
    'Número do Contrato','Data do Contrato','Situação','ID Cliente','Nome Cliente',
    'Unidade (ID)','Unidade (Nome)','Unidade (Número)','Valor Total do Contrato (R$)','Valor Saldo Aberto (R$)',
    'Valor Pago (R$)','ID Título (Recebível)','ID Título Cancelamento (a Pagar)'
  ];

  const matrix = [headers];
  rows.forEach(r => {
    matrix.push(headers.map(h => {
      const v = r[h];
      // para colunas financeiras mantemos number (empty string -> '')
      if ((h.indexOf('Valor') === 0 || h.indexOf('(R$)') !== -1) && (v !== '' && v !== null && v !== undefined)) return Number(v);
      // ID Cliente: manter number (inteiro) ou '' - já coerção foi feita antes
      if (h === 'ID Cliente') return (v === '' ? '' : Number(v));
      // Unidade (Número): manter number (inteiro) ou '' quando possível
      if (h === 'Unidade (Número)') return (v === '' ? '' : Number(v));
      // outros: deixar como estavam
      return v === undefined || v === null ? '' : v;
    }));
  });

  sheet.getRange(1,1,matrix.length,headers.length).setValues(matrix);

  // estilo cabeçalho
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1b5e20').setFontColor('white');
  sheet.setFrozenRows(1);

  // formata colunas monetárias
  const lastRow = sheet.getLastRow();
  const colIdx = headers.reduce((acc,h,i)=>{acc[h]=i+1;return acc;},{});
  const moneyCols = ['Valor Total do Contrato (R$)','Valor Saldo Aberto (R$)','Valor Pago (R$)'];
  moneyCols.forEach(h=>{
    const c = colIdx[h];
    if (c && lastRow > 1) {
      try { sheet.getRange(2, c, lastRow-1, 1).setNumberFormat('R$ #,##0.00'); } catch(e){/* ignore */ }
    }
  });

  // formatar ID Cliente como inteiro (coluna numérica)
  const idClienteCol = colIdx['ID Cliente'];
  if (idClienteCol && lastRow > 1) {
    try { sheet.getRange(2, idClienteCol, lastRow-1, 1).setNumberFormat('0'); } catch(e) { /* ignore */ }
  }

  // formatar 'Unidade (Número)' como inteiro quando houver dados
  const unidadeNumCol = colIdx['Unidade (Número)'];
  if (unidadeNumCol && lastRow > 1) {
    try { sheet.getRange(2, unidadeNumCol, lastRow-1, 1).setNumberFormat('0'); } catch(e) { /* ignore */ }
  }

  // para IDs/números que você queira como texto, liste aqui (REMOVA 'ID Cliente' se não quiser texto)
  const forceTextHeaders = [
    'ID Contrato','Número do Contrato','Unidade (ID)','ID Título (Recebível)','ID Título Cancelamento (a Pagar)'
  ];
  forceTextHeaders.forEach(h=>{
    const c = colIdx[h];
    if (c && lastRow > 1) {
      try { sheet.getRange(2, c, lastRow-1, 1).setNumberFormat('@'); } catch(e){/* ignore */ }
    }
  });

  try { sheet.autoResizeColumns(1, headers.length); } catch(e) {}
  console.log(`📝 ${rows.length} linhas escritas na aba "${CONFIG_CONTRATOS.SHEET_NAME}"`);
}

// ========== PRINCIPAL ==========
function atualizarContratosSienge() {
  try {
    console.log('🚀 Iniciando Contratos...');
    const items = fetchAllContracts_();
    const rows = processContractsToRows_(items);
    escreverContratosNaPlanilha_(rows);
    console.log('✅ Concluído atualizarContratosSienge');
  } catch (e) {
    console.error('❌ Erro em atualizarContratosSienge:', e);
  }
}
