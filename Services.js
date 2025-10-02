/*
 * Serviços – GAS_DevolucoesNFe
 * Requer as abas:
 *  Base:  ['ID','ChaveNFe','Numero','Serie','Emissao','Emitente_Nome','Emitente_CNPJ','Destinatario_Nome','Destinatario_CNPJ','CFOP','ValorNF','ValorICMS','XML_DriveURL','Status','CriadoPorCode','CriadoEm','ValidadoPorCode','ValidadorNome','ValidadoEm','FormaPagamento','DataPagamento','Anexo_DriveURL','Observacoes','AtualizadoEm']
 *  Itens: ['ID_RegistroBase','Seq','Codigo','Descricao','NCM','CFOP','Quantidade','VlrUnit','VlrTotal','StatusItem','MotivoItem','ObsItem','QtdeDevolvida','ValidadoPorCode_Item','ValidadoEm_Item']
 *  Users: ['PIN','NomeUsuario','Ativo']
 */
const PLANILHA_ID = '1htRCNJ79wg5Nh9XmnJ46_wDLUXjztwqRRJChJclS1rU';

const Services = (function () {

  // ------------- Utils -------------
  function getSpreadsheet() {
    return SpreadsheetApp.openById(PLANILHA_ID);
  }
  function idxByHeader_(headers) {
    const H = headers.map(h => String(h || '').trim());
    return [H, Object.fromEntries(H.map((h, i) => [h, i]))];
  }

  // ------------- Params -------------
  function getParam(param) {
    const cache = CacheService.getScriptCache();
    const key = 'param_' + param;
    let value = cache.get(key);
    if (value) return value;
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PARAMS);
    if (!sh) return '';
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(param)) {
        value = String(data[i][1]);
        cache.put(key, value, 300);
        return value;
      }
    }
    return '';
  }

  // ------------- Users / PIN -------------
  function ensureUsersSheet_(ss) {
    let sh = ss.getSheetByName(SHEET_USERS);
    if (!sh) {
      sh = ss.insertSheet(SHEET_USERS);
      sh.appendRow(['PIN','NomeUsuario','Ativo']);
    }
    return sh;
  }
  function validarPIN(pin) {
    if (!pin) return { ok:false, message:'PIN vazio' };
    const ss = getSpreadsheet();
    const sh = ensureUsersSheet_(ss);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(pin) && String(data[i][2]).toLowerCase() === 'true') {
        return { ok:true, pin, nome:data[i][1] };
      }
    }
    return { ok:false, message:'PIN inválido ou inativo' };
  }

  // ------------- Upload -------------
  function handleUpload(e, userCode, userEmail) {
    try {
      const xmls = [];
      if (e.postData && e.postData.type === 'application/json') {
        const body = JSON.parse(e.postData.contents || '{}');
        if (body.xml) xmls.push(body.xml);
      } else if (e.parameter && e.parameter.xml) {
        xmls.push(e.parameter.xml);
      }
      if (!xmls.length) return { ok:false, code:'NO_FILES', message:'Nenhum XML informado' };

      const created = [], errors = [];
      xmls.forEach(x => {
        try {
          const { base, itens } = parseXml_(x);
          const id = saveRecord_(base, itens, userCode);
          created.push({ id, chave: base.chaveNFe });
        } catch (err) {
          errors.push({ message: String(err && err.message || err) });
        }
      });
      return { ok:true, created, errors };
    } catch (err) {
      return { ok:false, code:'UPLOAD_ERROR', message:String(err && err.message || err) };
    }
  }

  // ------------- Listar -------------
  function listar(params, userCode) {
    try {
      const ss = getSpreadsheet();
      const sh = ss.getSheetByName(SHEET_BASE);
      if (!sh) return { ok:true, rows:[], page:1, pageSize:20, total:0 };

      const data = sh.getDataRange().getValues();
      if (data.length < 2) return { ok:true, rows:[], page:1, pageSize:20, total:0 };

      const headers = data[0];
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const r = {};
        headers.forEach((h, j) => {
          let v = data[i][j];
          if (v instanceof Date) v = v.toISOString();
          r[h] = v;
        });
        // ignora linha 100% vazia
        if (Object.values(r).some(v => v !== '' && v !== null && v !== undefined)) rows.push(r);
      }

      // --- Filtros ---
      const emitente = String(params && params.emitente || '').toLowerCase();
      const status   = String(params && params.status   || '').toLowerCase();
      const emiIni   = String(params && params.emissaoIni || '');
      const emiFim   = String(params && params.emissaoFim || '');

      const filtered = rows.filter(r => {
        // Restringe ao PIN do usuário
        if (userCode && String(r['CriadoPorCode'] || '') !== String(userCode)) return false;
        if (emitente && String(r['Emitente_Nome'] || '').toLowerCase().indexOf(emitente) === -1) return false;
        if (status && String(r['Status'] || '').toLowerCase() !== status) return false;
        if (emiIni || emiFim) {
          const raw = r['Emissao'];
          let iso;
          if (raw instanceof Date) {
            iso = raw.toISOString().slice(0,10);
          } else {
            try { 
              const dt = new Date(raw); 
              if (!isNaN(dt.getTime())) iso = dt.toISOString().slice(0,10); 
            } catch(e){}
          }
          if (!iso) return false;
          if (emiIni && iso < emiIni) return false;
          if (emiFim && iso > emiFim) return false;
        }
        return true;
      });

      const pageSize = params && params.pageSize ? parseInt(params.pageSize,10) : 20;
      const page     = params && params.page     ? parseInt(params.page,10)     : 1;
      const total    = filtered.length;
      const start    = Math.max(0, (page-1)*pageSize);
      const end      = Math.min(start + pageSize, total);
      
      // Calcula resumos para o cabeçalho
      const summary = {
        total: filtered.length,
        pendentes: filtered.filter(r => String(r['Status'] || '').toLowerCase() === 'pendente').length,
        recusadas: filtered.filter(r => String(r['Status'] || '').toLowerCase() === 'recusada').length
      };
      
      return { ok:true, rows: filtered.slice(start, end), page, pageSize, total, summary };
    } catch (err) {
      return { ok:false, code:'LISTAR_ERROR', message:String(err && err.message || err) };
    }
  }

 
 // ------------- Detalhar (super-robusto + debug) -------------
function detalhar(id, userCode) {
  try {
    if (!id) return { ok:false, code:'BAD_ID', message:'ID não informado' };

    const ss = getSpreadsheet();

    // --- BASE ---
    const baseSh = ss.getSheetByName(SHEET_BASE);
    if (!baseSh) return { ok:false, code:'NO_BASE', message:'Aba Base não encontrada' };
    const baseData = baseSh.getDataRange().getValues();
    if (!baseData.length) return { ok:false, code:'EMPTY_BASE', message:'Base vazia' };

    const baseHeaders = baseData[0].map(h => String(h||'').trim());
    const baseIdx     = Object.fromEntries(baseHeaders.map((h,i)=>[h,i]));

    // acha linha por ID; se não achar, tenta todas as colunas
    let rowIndex = -1;
    const idCol = (baseIdx['ID'] ?? baseIdx['Id'] ?? baseIdx['id'] ?? 0);
    for (let i=1;i<baseData.length;i++){
      if (String(baseData[i][idCol]) === String(id)) { rowIndex = i; break; }
    }
    if (rowIndex < 0) {
      outer:
      for (let i=1;i<baseData.length;i++){
        for (let j=0;j<baseHeaders.length;j++){
          if (String(baseData[i][j]) === String(id)) { rowIndex = i; break outer; }
        }
      }
    }
    if (rowIndex < 0) return { ok:false, code:'NOT_FOUND', message:'Registro não encontrado na Base' };

    const baseObj = {};
  baseHeaders.forEach((h, j) => {
  let v = baseData[rowIndex][j];
  if (v instanceof Date) v = v.toISOString();
  baseObj[h] = v;
});


    // --- ITENS ---
    const itensSh = ss.getSheetByName(SHEET_ITENS);
    let itens = [];
    let matchedBy = 'none';
    let itHeaders = [];
    let baseId = baseObj.ID || baseObj.Id || baseObj.id || id;
    
    if (itensSh) {
      const itData = itensSh.getDataRange().getValues();
      if (itData.length) {
        itHeaders = itData[0].map(h => String(h||'').trim());
        const itIdx = Object.fromEntries(itHeaders.map((h,i)=>[h,i]));

        const linkIdx =
          (itIdx['ID_RegistroBase'] != null) ? itIdx['ID_RegistroBase'] :
          (itIdx['ID']              != null) ? itIdx['ID']              :
          0;

        // 1) tenta vínculo direto
        for (let i=1;i<itData.length;i++){
          if (String(itData[i][linkIdx]) === String(baseId)) {
            const obj = {};
            itHeaders.forEach((h,j)=> obj[h] = itData[i][j]);
            itens.push(obj);
          }
        }
        if (itens.length) matchedBy = 'link';

        // 2) fallback: vasculha todas as colunas (se nada encontrado)
        if (itens.length === 0) {
          for (let i=1;i<itData.length;i++){
            for (let j=0;j<itHeaders.length;j++){
              if (String(itData[i][j]) === String(baseId)) {
                const obj = {};
                itHeaders.forEach((h,k)=> obj[h] = itData[i][k]);
                itens.push(obj);
                break;
              }
            }
          }
          if (itens.length) matchedBy = 'scan';
        }
      }
    }

    // Normaliza datas nos itens (após popular a lista)
    itens.forEach(item => {
      Object.keys(item).forEach(k => {
        if (item[k] instanceof Date) item[k] = item[k].toISOString();
      });
    });

    return {
      ok: true,
      base: baseObj,
      itens: itens,
      pastaDriveUrl: '',
      debug: { baseId, matchedBy, itHeaders }  // <— ajuda no console
    };

  } catch (err) {
    return { ok:false, code:'DETALHAR_ERROR', message:String(err && err.message || err) };
  }
}



  // ------------- Validar (Salvar) -------------
  function validar(payload, userCode, userEmail, userName) {
    try {
      const id = payload && payload.id;
      const status = payload && payload.status;
      if (!id || !status) return { ok:false, code:'MISSING_FIELDS', message:'ID ou status ausente' };

      const ss = getSpreadsheet();

      // BASE
      const baseSh = ss.getSheetByName(SHEET_BASE);
      const baseData = baseSh.getDataRange().getValues();
      const [BH, BIDX] = idxByHeader_(baseData[0]);

      const idCol = BIDX['ID'] ?? BIDX['Id'] ?? BIDX['id'] ?? 0;
      let row = -1;
      for (let i = 1; i < baseData.length; i++) {
        if (String(baseData[i][idCol]) === String(id)) { row = i; break; }
      }
      if (row < 0) return { ok:false, code:'NOT_FOUND', message:'Registro não encontrado' };

      const now = new Date();
      function setIfExists(colName, value) {
        if (BIDX[colName] != null) baseSh.getRange(row+1, BIDX[colName]+1).setValue(value);
      }
      setIfExists('Status', status);
      setIfExists('ValidadoPorCode', userCode);
      setIfExists('ValidadorNome', userName || '');
      setIfExists('ValidadoEm', now);
      if (status === 'Aceita') {
        setIfExists('FormaPagamento', payload.formaPagamento || '');
        setIfExists('DataPagamento',  payload.dataPagamento  || '');
        setIfExists('Anexo_DriveURL', payload.anexoUrl       || '');
      }
      setIfExists('Observacoes', payload.observacoes || '');
      setIfExists('AtualizadoEm', now);

      // ITENS
      const itensSh = ss.getSheetByName(SHEET_ITENS);
      if (itensSh) {
        const itData = itensSh.getDataRange().getValues();
        if (itData.length) {
          const [IH, IIDX] = idxByHeader_(itData[0]);
          const linkIdx = IIDX['ID_RegistroBase'] ?? IIDX['ID'] ?? 0;
          const seqIdx  = IIDX['Seq'] ?? IIDX['SEQ'] ?? 1;
          const upd = payload.itens || [];
          for (let i = 1; i < itData.length; i++) {
            if (String(itData[i][linkIdx]) === String(id)) {
              const seq = itData[i][seqIdx];
              const itemUpd = upd.find(u => parseInt(u.seq,10) === parseInt(seq,10));
              if (itemUpd) {
                function setI(col, val) { if (IIDX[col] != null) itensSh.getRange(i+1, IIDX[col]+1).setValue(val); }
                setI('StatusItem',           itemUpd.statusItem || '');
                setI('MotivoItem',           itemUpd.motivoItem || '');
                setI('QtdeDevolvida',        itemUpd.qtdeDevolvida || '');
                setI('ObsItem',              itemUpd.obsItem || '');
                setI('ValidadoPorCode_Item', userCode);
                setI('ValidadoEm_Item',      now);
              }
            }
          }
        }
      }

      logAction(userCode, userEmail, 'VALIDAR_' + status.toUpperCase(), id, '', 'Validação de nota');
      return { ok:true };
    } catch (err) {
      return { ok:false, code:'VALIDATE_ERROR', message:String(err && err.message || err) };
    }
  }

  // ------------- Parser + Save internos -------------
  function parseXml_(xmlStr) {
    const doc  = XmlService.parse(xmlStr);
    const root = doc.getRootElement();
    function getChildAnyNS(parent, name) {
      if (!parent) return null;
      return parent.getChild(name) || parent.getChild(name, parent.getNamespace()) || (function(){
        const list = parent.getChildren();
        for (let i=0;i<list.size?list.size():list.length;i++){
          const el = list.get ? list.get(i) : list[i];
          if (el && el.getName && el.getName() === name) return el;
        }
        return null;
      })();
    }
    function getTextAnyNS(parent, name) {
      const el = getChildAnyNS(parent, name);
      return el ? el.getText() : '';
    }

    let infNFe;
    if (root.getName() === 'nfeProc') {
      const nfe = getChildAnyNS(root, 'NFe'); infNFe = getChildAnyNS(nfe, 'infNFe');
    } else if (root.getName() === 'NFe') {
      infNFe = getChildAnyNS(root, 'infNFe');
    }
    if (!infNFe) throw new Error('XML NFe não reconhecido');

    const ide  = getChildAnyNS(infNFe, 'ide');
    const emit = getChildAnyNS(infNFe, 'emit');
    const dest = getChildAnyNS(infNFe, 'dest');
    const tot  = getChildAnyNS(infNFe, 'total');
    const icms = tot && getChildAnyNS(tot, 'ICMSTot');

    const idAttr = infNFe.getAttribute('Id');
    const chaveNFe = (idAttr && idAttr.getValue().replace(/^NFe/, '')) || '';
    const base = {
      chaveNFe: chaveNFe || 'não informado',
      numero: ide ? (getTextAnyNS(ide, 'nNF') || 'não informado') : 'não informado',
      serie:  ide ? (getTextAnyNS(ide, 'serie') || 'não informado') : 'não informado',
      emissao: ide ? (getTextAnyNS(ide, 'dhEmi') || getTextAnyNS(ide, 'dEmi') || 'não informado') : 'não informado',
      cfop: 'não informado',
      emitenteNome: emit ? (getTextAnyNS(emit, 'xNome') || 'não informado') : 'não informado',
      emitenteCNPJ: emit ? (getTextAnyNS(emit, 'CNPJ') || getTextAnyNS(emit, 'CPF') || 'não informado') : 'não informado',
      destinatarioNome: dest ? (getTextAnyNS(dest, 'xNome') || 'não informado') : 'não informado',
      destinatarioCNPJ: dest ? (getTextAnyNS(dest, 'CNPJ') || getTextAnyNS(dest, 'CPF') || 'não informado') : 'não informado',
      valorNF:   icms ? (getTextAnyNS(icms, 'vNF')   || '0') : '0',
      valorICMS: icms ? (getTextAnyNS(icms, 'vICMS') || '0') : '0'
    };

    const itens = [];
    const children = infNFe.getChildren();
    const len = children.size ? children.size() : children.length;
    for (let i=0;i<len;i++){
      const det = children.get ? children.get(i) : children[i];
      if (!det || det.getName() !== 'det') continue;
      const prod = getChildAnyNS(det, 'prod');
      const cfop = prod ? (getTextAnyNS(prod, 'CFOP') || '') : '';
      if (base.cfop === 'não informado' && cfop) base.cfop = cfop;
      const nItemAttr = det.getAttribute('nItem');
      const seqVal = nItemAttr ? parseInt(nItemAttr.getValue(),10) : i+1;
      itens.push({
        seq: seqVal,
        codigo:     prod ? (getTextAnyNS(prod, 'cProd') || 'não informado') : 'não informado',
        descricao:  prod ? (getTextAnyNS(prod, 'xProd') || 'não informado') : 'não informado',
        ncm:        prod ? (getTextAnyNS(prod, 'NCM')  || 'não informado') : 'não informado',
        cfop:       cfop || 'não informado',
        quantidade: prod ? (getTextAnyNS(prod, 'qCom')  || '0') : '0',
        vlrUnit:    prod ? (getTextAnyNS(prod, 'vUnCom')|| '0') : '0',
        vlrTotal:   prod ? (getTextAnyNS(prod, 'vProd') || '0') : '0'
      });
    }

    return { base, itens };
  }

  function saveRecord_(base, itens, userCode) {
    const ss = getSpreadsheet();
    const baseSh  = ss.getSheetByName(SHEET_BASE);
    const itensSh = ss.getSheetByName(SHEET_ITENS);

    // Verifica se já existe uma nota com a mesma chave
    const existingData = baseSh.getDataRange().getValues();
    if (existingData.length > 1) {
      const headers = existingData[0].map(h => String(h||'').trim());
      const chaveIdx = headers.indexOf('ChaveNFe');
      if (chaveIdx >= 0) {
        for (let i = 1; i < existingData.length; i++) {
          if (String(existingData[i][chaveIdx]) === String(base.chaveNFe)) {
            throw new Error(`Nota já existe na base: ${base.chaveNFe}`);
          }
        }
      }
    }

    const id = Utilities.getUuid();
    const folder = DriveApp.createFolder(id);
    const xmlUrl = folder.createFile('nota.xml', 'XML original não armazenado aqui.', MimeType.PLAIN_TEXT).getUrl();

    const now = new Date();
    baseSh.appendRow([
      id, base.chaveNFe, base.numero, base.serie, base.emissao,
      base.emitenteNome, base.emitenteCNPJ, base.destinatarioNome, base.destinatarioCNPJ,
      base.cfop, base.valorNF, base.valorICMS, xmlUrl,
      'Pendente', userCode, now, '', '', '', '', '', '', '', now
    ]);

    itens.forEach(it => {
      itensSh.appendRow([
        id, it.seq, it.codigo, it.descricao, it.ncm, it.cfop,
        it.quantidade, it.vlrUnit, it.vlrTotal,
        'Pendente', '', '', '', '', ''
      ]);
    });

    logAction(userCode, '', 'UPLOAD', id, '', 'Upload realizado');
    return id;
  }

  // ------------- Log -------------
  function logAction(userCode, userEmail, acao, idBase, seqItem, detalhes) {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(SHEET_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), userCode, userEmail, acao, idBase, seqItem || '', detalhes || '']);
  }



// Calcula métricas gerais para o dashboard
function dashboard(params) {
  try {
    const ss = getSpreadsheet();
    const baseSh = ss.getSheetByName(SHEET_BASE);
    if (!baseSh) return { ok: false, code: 'NO_BASE', message: 'Aba Base não encontrada' };

    const baseData = baseSh.getDataRange().getValues();
    if (baseData.length < 2) {
      return { ok: true, summaryStatus: {}, summaryValues: {}, dailyUploads: [], reasons: {}, productivity: {}, lastActivities: [] };
    }

    const headers = baseData[0].map(h => String(h || '').trim());
    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i);

    const summaryStatus = {};
    const summaryValues = {};
    const dailyUploads = {};
    const productivity = { created: {}, validated: {} };
    const baseRows = [];

    // Percorre cada linha da Base (começando em 1)
    for (let i = 1; i < baseData.length; i++) {
      const row = {};
      headers.forEach((h, j) => row[h] = baseData[i][j]);
      baseRows.push(row);

      const status = String(row['Status'] || '').trim();
      if (status) {
        summaryStatus[status] = (summaryStatus[status] || 0) + 1;
        const valor = parseFloat(String(row['ValorNF']).replace(',', '.')) || 0;
        summaryValues[status] = (summaryValues[status] || 0) + valor;
      }

      const criador = String(row['CriadoPorCode'] || '');
      if (criador) productivity.created[criador] = (productivity.created[criador] || 0) + 1;

      const validador = String(row['ValidadoPorCode'] || '');
      if (validador) productivity.validated[validador] = (productivity.validated[validador] || 0) + 1;

      // Contagem por dia de criação (CriadoEm)
      let dataCriacao = row['CriadoEm'];
      let isoDate;
      if (dataCriacao instanceof Date) {
        isoDate = dataCriacao.toISOString().slice(0, 10);
      } else {
        try {
          const dt = new Date(dataCriacao);
          if (!isNaN(dt.getTime())) isoDate = dt.toISOString().slice(0, 10);
        } catch (err) {}
      }
      if (isoDate) dailyUploads[isoDate] = (dailyUploads[isoDate] || 0) + 1;
    }

    // Contagem de motivos de recusa na aba Itens (StatusItem === 'Recusado')
    const itensSh = ss.getSheetByName(SHEET_ITENS);
    const reasons = {};
    if (itensSh) {
      const itData = itensSh.getDataRange().getValues();
      if (itData.length > 1) {
        const itHeaders = itData[0].map(h => String(h || '').trim());
        const itIdx = {};
        itHeaders.forEach((h, i) => itIdx[h] = i);

        for (let i = 1; i < itData.length; i++) {
          const statusItem = String(itData[i][itIdx['StatusItem']] || '').trim();
          const motivo = String(itData[i][itIdx['MotivoItem']] || '').trim();
          if (statusItem === 'Recusado' && motivo) {
            reasons[motivo] = (reasons[motivo] || 0) + 1;
          }
        }
      }
    }

    // Converte dailyUploads para array ordenado
    const dailyArray = Object.keys(dailyUploads)
      .sort()
      .map(d => ({ date: d, count: dailyUploads[d] }));

    // Converte produtividade para arrays
    const createdArray = Object.keys(productivity.created).map(u => ({ user: u, count: productivity.created[u] }));
    const validatedArray = Object.keys(productivity.validated).map(u => ({ user: u, count: productivity.validated[u] }));

    // Últimas atividades (últimos 5 registros ordenados por CriadoEm decrescente)
    baseRows.sort((a, b) => {
      const ta = a['CriadoEm'] instanceof Date ? a['CriadoEm'].getTime() : new Date(a['CriadoEm']).getTime();
      const tb = b['CriadoEm'] instanceof Date ? b['CriadoEm'].getTime() : new Date(b['CriadoEm']).getTime();
      return tb - ta;
    });
    const lastActivities = baseRows.slice(0, 5).map(row => ({
      id: row['ID'] || '',
      chave: row['ChaveNFe'] || '',
      emissao: row['Emissao'] instanceof Date ? row['Emissao'].toISOString() : String(row['Emissao'] || ''),
      status: row['Status'] || '',
      valor: parseFloat(String(row['ValorNF']).replace(',', '.')) || 0,
      criadoEm: row['CriadoEm'] instanceof Date ? row['CriadoEm'].toISOString() : String(row['CriadoEm'] || ''),
      criadoPor: String(row['CriadoPorCode'] || '')
    }));

    return {
      ok: true,
      summaryStatus,
      summaryValues,
      dailyUploads: dailyArray,
      reasons,
      productivity: { created: createdArray, validated: validatedArray },
      lastActivities
    };
  } catch (err) {
    return { ok: false, code: 'DASHBOARD_ERROR', message: String(err) };
  }
}

// Dashboard filtrado para o usuário logado
function dashboardPessoal(userCode, userName) {
  try {
    const ss = getSpreadsheet();
    const baseSh = ss.getSheetByName(SHEET_BASE);
    if (!baseSh) return { ok: false, code: 'NO_BASE', message: 'Aba Base não encontrada' };

    const baseData = baseSh.getDataRange().getValues();
    if (baseData.length < 2) {
      return { ok: true, summaryStatus: {}, summaryValues: {}, dailyUploads: [], reasons: {}, productivity: {}, lastActivities: [] };
    }

    const headers = baseData[0].map(h => String(h || '').trim());
    const colIndex = {};
    headers.forEach((h, i) => colIndex[h] = i);

    const summaryStatus = {};
    const summaryValues = {};
    const dailyUploads = {};
    const baseRows = [];

    // Filtra linhas criadas pelo usuário
    for (let i = 1; i < baseData.length; i++) {
      const row = {};
      headers.forEach((h, j) => row[h] = baseData[i][j]);

      if (String(row['CriadoPorCode'] || '') !== String(userCode)) continue;
      baseRows.push(row);

      const status = String(row['Status'] || '').trim();
      if (status) {
        summaryStatus[status] = (summaryStatus[status] || 0) + 1;
        const valor = parseFloat(String(row['ValorNF']).replace(',', '.')) || 0;
        summaryValues[status] = (summaryValues[status] || 0) + valor;
      }

      // Contagem por dia (CriadoEm)
      let dataCriacao = row['CriadoEm'];
      let isoDate;
      if (dataCriacao instanceof Date) {
        isoDate = dataCriacao.toISOString().slice(0, 10);
      } else {
        try {
          const dt = new Date(dataCriacao);
          if (!isNaN(dt.getTime())) isoDate = dt.toISOString().slice(0, 10);
        } catch (err) {}
      }
      if (isoDate) dailyUploads[isoDate] = (dailyUploads[isoDate] || 0) + 1;
    }

    // Motivos de recusa apenas para itens vinculados às notas do usuário
    const itensSh = ss.getSheetByName(SHEET_ITENS);
    const reasons = {};
    if (itensSh && baseRows.length > 0) {
      const itData = itensSh.getDataRange().getValues();
      const itHeaders = itData[0].map(h => String(h || '').trim());
      const itIdx = {};
      itHeaders.forEach((h, i) => itIdx[h] = i);

      // Cria um conjunto de IDs da base do usuário
      const idsUsuario = {};
      baseRows.forEach(r => { idsUsuario[String(r['ID'])] = true; });

      for (let i = 1; i < itData.length; i++) {
        const idBase = String(itData[i][itIdx['ID_RegistroBase']] || '');
        if (!idsUsuario[idBase]) continue;
        const statusItem = String(itData[i][itIdx['StatusItem']] || '').trim();
        const motivo = String(itData[i][itIdx['MotivoItem']] || '').trim();
        if (statusItem === 'Recusado' && motivo) {
          reasons[motivo] = (reasons[motivo] || 0) + 1;
        }
      }
    }

    // Converte dailyUploads para array ordenada
    const dailyArray = Object.keys(dailyUploads)
      .sort()
      .map(d => ({ date: d, count: dailyUploads[d] }));

    // Últimas atividades: ordena por CriadoEm decrescente (limita a 5)
    baseRows.sort((a, b) => {
      const ta = a['CriadoEm'] instanceof Date ? a['CriadoEm'].getTime() : new Date(a['CriadoEm']).getTime();
      const tb = b['CriadoEm'] instanceof Date ? b['CriadoEm'].getTime() : new Date(b['CriadoEm']).getTime();
      return tb - ta;
    });
    const lastActivities = baseRows.slice(0, 5).map(row => ({
      id: row['ID'] || '',
      chave: row['ChaveNFe'] || '',
      emissao: row['Emissao'] instanceof Date ? row['Emissao'].toISOString() : String(row['Emissao'] || ''),
      status: row['Status'] || '',
      valor: parseFloat(String(row['ValorNF']).replace(',', '.')) || 0,
      criadoEm: row['CriadoEm'] instanceof Date ? row['CriadoEm'].toISOString() : String(row['CriadoEm'] || ''),
      criadoPor: String(row['CriadoPorCode'] || '')
    }));

    return {
      ok: true,
      summaryStatus,
      summaryValues,
      dailyUploads: dailyArray,
      reasons,
      productivity: {}, // neste caso, produtividade individual não é relevante
      lastActivities
    };
  } catch (err) {
    return { ok: false, code: 'DASHBOARD_PESSOAL_ERROR', message: String(err) };
  }
}


  // ------------- API pública -------------
  return {
    getParam,
    validarPIN,
    handleUpload,
    listar,
    detalhar,
    validar,
      dashboard,     
  dashboardPessoal,  
    logAction
  };
})();
