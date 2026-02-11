/**
 * SIENGE · TÍTULOS INADIMPLENTES (1 companyId por execução)
 * Endpoint: /public/api/bulk-data/v1/defaulters-receivable-bills
 * Aba: "Titulos inadimplentes"
 */

const CONFIG_INAD = {
  // Credenciais (Script Properties)
  SIENGE_SUBDOMAIN: PropertiesService.getScriptProperties().getProperty('SIENGE_SUBDOMAIN'),
  SIENGE_USER:      PropertiesService.getScriptProperties().getProperty('SIENGE_USER'),
  SIENGE_PASS:      PropertiesService.getScriptProperties().getProperty('SIENGE_PASS'),

  // Filtros
  COMPANY_ID: 2007,               // obrigatório por execução
  ENTERPRISE_ID: 78,           // empreend. (ou null para não filtrar)
  SHOW_ONLY_DEFAULTERS: false, // mantém compatibilidade - mas veja abaixo

  // Flags extras (ajudam a trazer FI e parciais)
  INCLUDE_PARTIALLY_PAID: true,        // includePartiallyPaidInstallments
  DEFAULTERS_RECEIVABLE_BILLS: true,   // defaultersReceivableBills (agora true por padrão)
  IN_BILLING_RECEIVABLE_BILLS: true,   // inBillingReceivableBills (agora true por padrão)

  // Planilha
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  SHEET_NAME: 'Titulos inadimplentes'
};

// ===== Helpers =====
function _toNumberBR(v){
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isNaN(v) ? '' : v;
  let s = String(v).trim();
  if (!s) return '';
  s = s.replace(/[^\d.,-]/g,'');
  if (s === '') return '';
  if (s.indexOf('.') > -1 && s.indexOf(',') > -1) { s = s.replace(/\./g,'').replace(',','.'); }
  else if (s.indexOf(',') > -1 && s.indexOf('.') === -1) { s = s.replace(',','.'); }
  const n = Number(s);
  return isNaN(n) ? '' : n;
}
function _siglaDoDocumento(doc){
  const s = (doc || '').toString();
  if (/INCC/i.test(s)) return 'INCC';
  if (/RCCT/i.test(s)) return 'RCCT';
  return 'CT';
}
function _unitLabel(u){
  if (!u) return '';
  if (typeof u === 'object') {
    const id = u.id ?? u.unitId ?? u.code ?? u.number ?? u.identifier ?? u.uid ?? '';
    const nome = u.name ?? u.unitName ?? u.description ?? u.title ?? u.label ?? '';
    const combo = [id, nome].filter(Boolean).join(' - ');
    return combo || JSON.stringify(u);
  }
  return String(u);
}

// ===== Main =====
function atualizarInadimplentes(){
  try{
    console.log(`🚀 Início Inadimplentes (companyId=${CONFIG_INAD.COMPANY_ID}, enterpriseId=${CONFIG_INAD.ENTERPRISE_ID})`);
    const dados = buscarInadimplentesAPI_();
    const linhas = processarInadimplentes_(dados);
    // Ordena por ID Cliente (numericamente). IDs vazios vão pro final.
    linhas.sort((a,b)=>{
      const A = (a['ID Cliente'] === '' || a['ID Cliente'] === null || a['ID Cliente'] === undefined) ? Infinity : Number(a['ID Cliente']);
      const B = (b['ID Cliente'] === '' || b['ID Cliente'] === null || b['ID Cliente'] === undefined) ? Infinity : Number(b['ID Cliente']);
      return A - B;
    });
    escreverNaPlanilhaInad_(linhas);
    console.log(`✅ Concluído: ${linhas.length} linhas em "${CONFIG_INAD.SHEET_NAME}"`);
  }catch(e){
    console.error('❌ Erro:', e);
  }
}

// ===== Fetch =====
function buscarInadimplentesAPI_(){
  const base = `https://api.sienge.com.br/${CONFIG_INAD.SIENGE_SUBDOMAIN}/public/api/bulk-data/v1/defaulters-receivable-bills`;
  // Monta query string de forma clara e robusta
  const parts = [];
  parts.push(`companyId=${encodeURIComponent(CONFIG_INAD.COMPANY_ID)}`);
  parts.push(`showOnlyDefaulters=${CONFIG_INAD.SHOW_ONLY_DEFAULTERS ? 'true' : 'false'}`);

  if (CONFIG_INAD.ENTERPRISE_ID !== null && CONFIG_INAD.ENTERPRISE_ID !== undefined && CONFIG_INAD.ENTERPRISE_ID !== '') {
    parts.push(`enterpriseId=${encodeURIComponent(CONFIG_INAD.ENTERPRISE_ID)}`);
  }
  if (CONFIG_INAD.INCLUDE_PARTIALLY_PAID) parts.push('includePartiallyPaidInstallments=true');

  // --- OS DOIS PARAMS QUE VOCÊ QUER ATIVAR ---
  if (CONFIG_INAD.DEFAULTERS_RECEIVABLE_BILLS) parts.push('defaultersReceivableBills=true');
  if (CONFIG_INAD.IN_BILLING_RECEIVABLE_BILLS) parts.push('inBillingReceivableBills=true');
  // ------------------------------------------------

  const qs = parts.join('&');
  const url = `${base}?${qs}`;

  const options = {
    method:'GET',
    headers:{
      Authorization:'Basic '+Utilities.base64Encode(`${CONFIG_INAD.SIENGE_USER}:${CONFIG_INAD.SIENGE_PASS}`),
      Accept:'application/json'
    },
    muteHttpExceptions:true,
    timeout: 60000
  };

  console.log(`📡 GET: ${url}`);
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  if (code !== 200){
    throw new Error(`HTTP ${code} :: ${resp.getContentText()}`);
  }

  const json = JSON.parse(resp.getContentText());
  const data = Array.isArray(json?.data) ? json.data : [];
  console.log(`📦 Recebidos: ${data.length} registros`);

  // Estatística por condição (opcional)
  const contPorCond = {};
  data.forEach(r=>{
    (r.defaulterInstallments || []).forEach(p=>{
      const c = (p.conditionType || '').toString().trim() || '(vazio)';
      contPorCond[c] = (contPorCond[c] || 0) + 1;
    });
  });
  console.log('📊 Parcelas por condição:', JSON.stringify(contPorCond));
  return data;
}

// ===== Process =====
function processarInadimplentes_(dados){
  const linhas = [];
  dados.forEach(reg=>{
    const parcelas = Array.isArray(reg.defaulterInstallments) && reg.defaulterInstallments.length ? reg.defaulterInstallments : [null];

    parcelas.forEach(par=>{
      const companyId = reg.companyId ?? reg.company?.id ?? '';
      // Coerção do ID do cliente para número inteiro quando possível
      const rawClientId = reg.clientId ?? reg.customer?.id ?? reg.customerId ?? '';
      const clientIdNum = _toNumberBR(rawClientId);
      const clientIdFinal = (clientIdNum === '') ? '' : Math.trunc(clientIdNum);
      const clientName = reg.clientName ?? reg.customer?.name ?? '';

      const receivableBillId = reg.receivableBillId ?? reg.billId ?? '';
      const issueDate = reg.issueDate ?? reg.emissionDate ?? '';
      const documentNumber = reg.documentNumber ?? reg.document ?? '';
      const siglaDocumento = _siglaDoDocumento(documentNumber);

      let costCentersId = '';
      if (Array.isArray(reg.costCentersId)) { costCentersId = reg.costCentersId.join(','); }
      else if (reg.costCenter && typeof reg.costCenter === 'object') { costCentersId = [reg.costCenter.id].filter(Boolean).join(','); }
      else { costCentersId = reg.costCentersId ?? ''; }

      let unitsStr = '';
      if (Array.isArray(reg.units)) { unitsStr = reg.units.map(_unitLabel).join(' | '); }
      else if (reg.unit) { unitsStr = _unitLabel(reg.unit); }
      else if (typeof reg.units === 'object' && reg.units !== null) { unitsStr = _unitLabel(reg.units); }
      else if (typeof reg.units === 'string' || typeof reg.units === 'number') { unitsStr = String(reg.units).trim(); }
      if (!unitsStr){
        const aliases = ['unitCode','unitName','unitNumber','unitId','unity','apartment','apto','block','tower','building','bloco','torre','unidade'];
        for (const k of aliases){ if (reg[k]){ unitsStr = String(reg[k]).trim(); break; } }
      }

      const receivableBillValue = _toNumberBR(reg.receivableBillValue ?? reg.originalValue ?? '');

      const linha = {
        'ID Empresa': companyId,
        // ID Cliente agora vem como número inteiro (ou '')
        'ID Cliente': clientIdFinal,
        'Nome Cliente': clientName,
        'ID Título': receivableBillId,
        'Data de Emissão': issueDate,
        'Número do Documento': documentNumber,
        'Sigla do Documento': siglaDocumento,
        'Centros de Custo (IDs)': costCentersId,
        'Unidades (IDs/Nomes)': unitsStr,
        'Valor do Título': receivableBillValue,

        'Parcela - ID': par?.installmentId ?? par?.id ?? '',
        'Parcela - Condição': par?.conditionType ?? '',
        'Parcela - Vencimento': par?.dueDate ?? '',
        'Parcela - Dias de Atraso': par?.daysOfDelay ?? '',

        'Parcela - Valor Corrigido s/ Acrésc.': _toNumberBR(par?.correctedValueWithoutAdditions ?? ''),
        'Parcela - Pro Rata': _toNumberBR(par?.proRata ?? ''),
        'Parcela - Juros': _toNumberBR(par?.interest ?? ''),
        'Parcela - Multa': _toNumberBR(par?.fine ?? ''),
        'Parcela - Acréscimos Totais': _toNumberBR(par?.totalAdditions ?? ''),
        'Parcela - Valor Corrigido c/ Acrésc.': _toNumberBR(par?.correctedValueWithAdditions ?? ''),
        'Parcela - Nº': (par?.installmentNumber === null || par?.installmentNumber === undefined) ? '' : String(par.installmentNumber),
        'Parcela - Enviado SPC/Serasa': par?.installmentSentToSPCSerasa ?? '',
        'Parcela - Valor Numérico': _toNumberBR(par?.correctedValueWithAdditions ?? '')
      };

      linhas.push(linha);
    });
  });

  console.log(`✅ Linhas geradas: ${linhas.length}`);
  return linhas;
}

// ===== Write =====
function escreverNaPlanilhaInad_(dados){
  const ss = SpreadsheetApp.openById(CONFIG_INAD.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG_INAD.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_INAD.SHEET_NAME);
  sheet.clear();

  if (!dados.length){
    sheet.getRange(1,1).setValue('Sem dados');
    console.log('⚠️ Nenhum dado para escrever.');
    return;
  }

  const headers = [
    'ID Empresa','ID Cliente','Nome Cliente','ID Título','Data de Emissão',
    'Número do Documento','Sigla do Documento',
    'Centros de Custo (IDs)','Unidades (IDs/Nomes)','Valor do Título',
    'Parcela - ID','Parcela - Condição','Parcela - Vencimento','Parcela - Dias de Atraso',
    'Parcela - Valor Corrigido s/ Acrésc.','Parcela - Pro Rata','Parcela - Juros','Parcela - Multa',
    'Parcela - Acréscimos Totais','Parcela - Valor Corrigido c/ Acrésc.','Parcela - Nº',
    'Parcela - Enviado SPC/Serasa','Parcela - Valor Numérico'
  ];

  const values = [headers, ...dados.map(o => headers.map(h => {
    const v = o[h];
    // tratar colunas financeiras como número (se não vazio)
    if (['Valor do Título','Parcela - Valor Corrigido s/ Acrésc.','Parcela - Pro Rata','Parcela - Juros','Parcela - Multa','Parcela - Acréscimos Totais','Parcela - Valor Corrigido c/ Acrésc.','Parcela - Valor Numérico'].includes(h)) {
      return (v === '' || v === null || v === undefined) ? '' : Number(v);
    }
    // ID Cliente já é número inteiro ou ''
    if (h === 'ID Cliente') {
      return (v === '' || v === null || v === undefined) ? '' : Number(v);
    }
    return v === undefined || v === null ? '' : v;
  }))];

  sheet.getRange(1,1,values.length,headers.length).setValues(values);

  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);

  const money = new Set(['Valor do Título','Parcela - Valor Corrigido s/ Acrésc.','Parcela - Pro Rata','Parcela - Juros','Parcela - Multa','Parcela - Acréscimos Totais','Parcela - Valor Corrigido c/ Acrésc.','Parcela - Valor Numérico']);
  const last = sheet.getLastRow();
  headers.forEach((h,i)=>{
    const col = i+1;
    if(last>1){
      if(money.has(h)) {
        try { sheet.getRange(2,col,last-1,1).setNumberFormat('R$ #,##0.00'); } catch(e){/*ignore*/ }
      }
      if(h==='Parcela - Dias de Atraso') {
        try { sheet.getRange(2,col,last-1,1).setNumberFormat('0'); } catch(e){/*ignore*/ }
      }
      if(h==='Parcela - Nº') {
        try { sheet.getRange(2,col,last-1,1).setNumberFormat('@'); } catch(e){/*ignore*/ }
      }
    }
  });

  // Formata ID Cliente como inteiro (0)
  try {
    const idx = headers.indexOf('ID Cliente');
    if (idx !== -1 && last > 1) {
      sheet.getRange(2, idx+1, last-1, 1).setNumberFormat('0');
    }
  } catch(e){/*ignore*/}

  try { sheet.autoResizeColumns(1, headers.length); } catch(_){}
}
