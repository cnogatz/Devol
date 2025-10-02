/*
 * Serviços – GAS_DevolucoesNFe
 * Requer as abas:
 *  Base:  ['ID','ChaveNFe','Numero','Serie','Emissao','Emitente_Nome','Emitente_CNPJ','Destinatario_Nome','Destinatario_CNPJ','CFOP','ValorNF','ValorICMS','XML_DriveURL','Status','CriadoPorCode','CriadoEm','ValidadoPorCode','ValidadorNome','ValidadoEm','FormaPagamento','DataPagamento','Anexo_DriveURL','Observacoes','AtualizadoEm']
 *  Itens: ['ID_RegistroBase','Seq','Codigo','Descricao','NCM','CFOP','Quantidade','VlrUnit','VlrTotal','StatusItem','MotivoItem','ObsItem','QtdeDevolvida','ValidadoPorCode_Item','ValidadoEm_Item']
 *  Users: ['PIN','NomeUsuario','Ativo']
 */
const PLANILHA_ID = '1htRCNJ79wg5Nh9XmnJ46_wDLUXjztwqRRJChJclS1rU';

const BASE_HEADERS = [
  'ID','ChaveNFe','Numero','Serie','Emissao',
  'Emitente_Nome','Emitente_CNPJ','Destinatario_Nome','Destinatario_CNPJ',
  'CFOP','ValorNF','ValorICMS','XML_DriveURL','Status',
  'CriadoPorCode','CriadoEm','ValidadoPorCode','ValidadorNome','ValidadoEm',
  'FormaPagamento','DataPagamento','Anexo_DriveURL','Observacoes','AtualizadoEm'
];

const ITENS_HEADERS = [
  'ID_RegistroBase','Seq','Codigo','Descricao','NCM','CFOP','Quantidade','VlrUnit','VlrTotal',
  'StatusItem','MotivoItem','ObsItem','QtdeDevolvida','ValidadoPorCode_Item','ValidadoEm_Item'
];

const LOG_HEADERS = [
  'Timestamp','UserCode','UsuarioEmail','Acao','ID_RegistroBase','SeqItem','Detalhes'
];

const Services = (function () {

  // ------------- Utils -------------
  function getSpreadsheet() {
    return SpreadsheetApp.openById(PLANILHA_ID);
  }
  function idxByHeader_(headers) {
    const H = headers.map(h => String(h || '').trim());
    return [H, Object.fromEntries(H.map((h, i) => [h, i]))];
  }

  function ensureSheetWithHeaders_(ss, sheetName, headers) {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
    }
    if (headers && headers.length) {
      if (sh.getMaxColumns() < headers.length) {
        sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
      }
      const lastRow = sh.getLastRow();
      if (lastRow === 0) {
        sh.appendRow(headers);
      } else if (lastRow === 1) {
        const firstRowRange = sh.getRange(1, 1, 1, headers.length);
        const values = firstRowRange.getValues()[0];
        const hasAnyValue = values.some(v => String(v || '').trim() !== '');
        if (!hasAnyValue) {
          firstRowRange.setValues([headers]);
        }
      }
    }
    return sh;
  }

  function toNumber_(value) {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') return value;
    if (value instanceof Date) return value.getTime();
    const raw = String(value).trim().replace(/\s+/g, '');
    if (!raw) return 0;
    if (raw.includes('.') && raw.includes(',')) {
      return parseFloat(raw.replace(/\./g, '').replace(',', '.')) || 0;
    }
    if (raw.includes(',')) {
      return parseFloat(raw.replace(',', '.')) || 0;
    }
    return parseFloat(raw) || 0;
  }

  function parseDate_(value) {
    if (!value) return null;
    if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
    if (typeof value === 'number') {
      const d = new Date(value);
      return isNaN(d.getTime()) ? null : d;
    }
    const str = String(value).trim();
    if (!str) return null;
    let d = new Date(str);
    if (!isNaN(d.getTime())) return d;
    d = new Date(str.replace(' ', 'T'));
    if (!isNaN(d.getTime())) return d;
    const parts = str.split(/[\/]/);
    if (parts.length === 3) {
      const [dStr, mStr, yStr] = parts;
      const dd = parseInt(dStr, 10);
      const mm = parseInt(mStr, 10) - 1;
      const yy = parseInt(yStr, 10);
      if (!isNaN(dd) && !isNaN(mm) && !isNaN(yy)) {
        const alt = new Date(yy, mm, dd);
        return isNaN(alt.getTime()) ? null : alt;
      }
    }
    return null;
  }

  function monthKey_(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return { key: 'Sem data', label: 'Sem data' };
    }
    const y = date.getFullYear();
    const m = date.getMonth();
    const key = y + '-' + String(m + 1).padStart(2, '0');
    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    return { key, label: meses[m] + '/' + y };
  }

  function collectBaseRecords_() {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(SHEET_BASE);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0].map(h => String(h || '').trim());
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      let hasContent = false;
      headers.forEach((h, idx) => {
        const val = row[idx];
        if (val !== '' && val != null) hasContent = true;
        obj[h] = val;
      });
      if (hasContent) rows.push(obj);
    }
    return rows;
  }

  function filterRecordsByUser_(records, userCode) {
    const code = String(userCode || '').trim();
    if (!code) return [];
    return records.filter(rec => {
      const criado = String(rec.CriadoPorCode || '').trim();
      const validado = String(rec.ValidadoPorCode || '').trim();
      return criado === code || validado === code;
    });
  }

  function buildDashboardFromRecords_(records) {
    const statusMap = new Map();
    const valorStatusMap = new Map();
    const mesMap = new Map();
    const usuarioMap = new Map();
    let totalValor = 0;

    records.forEach(rec => {
      const status = String(rec.Status || 'Sem status');
      const valor = toNumber_(rec.ValorNF);
      totalValor += valor;

      statusMap.set(status, (statusMap.get(status) || 0) + 1);
      valorStatusMap.set(status, (valorStatusMap.get(status) || 0) + valor);

      const dataRef = parseDate_(rec.Emissao || rec.CriadoEm || rec.ValidadoEm);
      const mk = monthKey_(dataRef);
      const mesEntry = mesMap.get(mk.key) || { label: mk.label, count: 0, valor: 0 };
      mesEntry.count += 1;
      mesEntry.valor += valor;
      mesMap.set(mk.key, mesEntry);

      const responsavel = String(rec.ValidadorNome || rec.ValidadoPorCode || rec.CriadoPorCode || 'Sem responsável').trim() || 'Sem responsável';
      usuarioMap.set(responsavel, (usuarioMap.get(responsavel) || 0) + 1);
    });

    const resumo = {
      totalNotas: records.length,
      totalValor: Math.round(totalValor * 100) / 100,
      pendentes: statusMap.get('Pendente') || 0,
      aceitas: statusMap.get('Aceita') || 0,
      recusadas: statusMap.get('Recusada') || 0,
      ultimaAtualizacao: new Date()
    };

    function dataset(columns, rows) {
      return { columns, rows };
    }

    const statusRows = Array.from(statusMap.entries()).sort((a, b) => b[1] - a[1]).map(([s, q]) => [s, q]);
    const valorRows = Array.from(valorStatusMap.entries()).sort((a, b) => b[1] - a[1]).map(([s, v]) => [s, Math.round(v * 100) / 100]);

    const mesOrder = Array.from(mesMap.entries()).sort((a, b) => {
      if (a[0] === 'Sem data') return 1;
      if (b[0] === 'Sem data') return -1;
      return a[0].localeCompare(b[0]);
    });
    const mesRows = mesOrder.map(([, info]) => [info.label, info.count, Math.round(info.valor * 100) / 100]);

    const usuarioRows = Array.from(usuarioMap.entries()).sort((a, b) => b[1] - a[1]).map(([u, q]) => [u, q]);

    return {
      resumo,
      charts: {
        status: dataset([['string', 'Status'], ['number', 'Notas']], statusRows),
        valorPorStatus: dataset([['string', 'Status'], ['number', 'Valor NF']], valorRows),
        porMes: dataset([['string', 'Mês'], ['number', 'Notas'], ['number', 'Valor NF']], mesRows),
        porUsuario: dataset([['string', 'Responsável'], ['number', 'Notas']], usuarioRows)
      }
    };
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
    }
    if (sh.getLastRow() === 0) {
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

      const pageSize = params && params.pageSize ? parseInt(params.pageSize,10) : 20;
      const page     = params && params.page     ? parseInt(params.page,10)     : 1;
      const total    = rows.length;
      const start    = Math.max(0, (page-1)*pageSize);
      const end      = Math.min(start + pageSize, total);
      return { ok:true, rows: rows.slice(start, end), page, pageSize, total };
    } catch (err) {
      return { ok:false, code:'LISTAR_ERROR', message:String(err && err.message || err) };
    }
  }

  // ------------- Detalhar -------------
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
    baseHeaders.forEach((h,j)=> baseObj[h] = baseData[rowIndex][j]);

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


  function dashboard(params) {
    try {
      const registros = collectBaseRecords_();
      const dash = buildDashboardFromRecords_(registros);
      return Object.assign({ ok: true }, dash);
    } catch (err) {
      return { ok:false, code:'DASHBOARD_ERROR', message:String(err && err.message || err) };
    }
  }

  function dashboardPessoal(userCode, userName) {
    try {
      const registros = collectBaseRecords_();
      const meus = filterRecordsByUser_(registros, userCode);
      const dash = buildDashboardFromRecords_(meus);
      dash.resumo.usuario = userName || '';
      return Object.assign({ ok: true }, dash);
    } catch (err) {
      return { ok:false, code:'DASHBOARD_PESSOAL_ERROR', message:String(err && err.message || err) };
    }
  }

  function minhasAtividades(params, userCode) {
    try {
      const ss = getSpreadsheet();
      const sh = ss.getSheetByName(SHEET_LOG);
      if (!sh) return { ok:true, rows:[], total:0 };
      const data = sh.getDataRange().getValues();
      if (data.length <= 1) return { ok:true, rows:[], total:0 };
      const headers = data[0].map(h => String(h || '').trim());
      const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const user = String(row[idx['UserCode'] != null ? idx['UserCode'] : 1] || '').trim();
        if (user === String(userCode || '').trim()) {
          const obj = {};
          headers.forEach((h, j) => obj[h] = row[j]);
          rows.push(obj);
        }
      }
      const type = params && params.type;
      let filtered = rows;
      if (type === 'uploads') {
        filtered = rows.filter(r => String(r.Acao || '').toUpperCase().indexOf('UPLOAD') !== -1);
      } else if (type === 'validacoes') {
        filtered = rows.filter(r => String(r.Acao || '').toUpperCase().indexOf('VALIDAR') !== -1);
      }
      return { ok:true, rows: filtered, total: filtered.length };
    } catch (err) {
      return { ok:false, code:'ATIVIDADES_ERROR', message:String(err && err.message || err) };
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
    let infNFe;
    if (root.getName() === 'nfeProc') {
      const nfe = root.getChild('NFe'); infNFe = nfe && nfe.getChild('infNFe');
    } else if (root.getName() === 'NFe') {
      infNFe = root.getChild('infNFe');
    }
    if (!infNFe) throw new Error('XML NFe não reconhecido');

    const ide  = infNFe.getChild('ide');
    const emit = infNFe.getChild('emit');
    const dest = infNFe.getChild('dest');
    const tot  = infNFe.getChild('total');
    const icms = tot && tot.getChild('ICMSTot');

    const chaveNFe = (infNFe.getAttribute('Id') && infNFe.getAttribute('Id').getValue().replace(/^NFe/, '')) || '';
    const base = {
      chaveNFe: chaveNFe || 'não informado',
      numero: ide ? (ide.getChildText('nNF') || 'não informado') : 'não informado',
      serie:  ide ? (ide.getChildText('serie') || 'não informado') : 'não informado',
      emissao: ide ? (ide.getChildText('dhEmi') || ide.getChildText('dEmi') || 'não informado') : 'não informado',
      cfop: 'não informado',
      emitenteNome: emit ? (emit.getChildText('xNome') || 'não informado') : 'não informado',
      emitenteCNPJ: emit ? (emit.getChildText('CNPJ') || emit.getChildText('CPF') || 'não informado') : 'não informado',
      destinatarioNome: dest ? (dest.getChildText('xNome') || 'não informado') : 'não informado',
      destinatarioCNPJ: dest ? (dest.getChildText('CNPJ') || dest.getChildText('CPF') || 'não informado') : 'não informado',
      valorNF:   icms ? (icms.getChildText('vNF')   || '0') : '0',
      valorICMS: icms ? (icms.getChildText('vICMS') || '0') : '0'
    };

    const itens = [];
    (infNFe.getChildren('det') || []).forEach(det => {
      const prod = det.getChild('prod');
      const cfop = prod ? (prod.getChildText('CFOP') || '') : '';
      if (base.cfop === 'não informado' && cfop) base.cfop = cfop;
      itens.push({
        seq: parseInt(det.getAttribute('nItem').getValue(), 10),
        codigo:     prod ? (prod.getChildText('cProd') || 'não informado') : 'não informado',
        descricao:  prod ? (prod.getChildText('xProd') || 'não informado') : 'não informado',
        ncm:        prod ? (prod.getChildText('NCM')  || 'não informado') : 'não informado',
        cfop:       cfop || 'não informado',
        quantidade: prod ? (prod.getChildText('qCom')  || '0') : '0',
        vlrUnit:    prod ? (prod.getChildText('vUnCom')|| '0') : '0',
        vlrTotal:   prod ? (prod.getChildText('vProd') || '0') : '0'
      });
    });

    return { base, itens };
  }

  function saveRecord_(base, itens, userCode) {
    const ss = getSpreadsheet();
    const baseSh  = ensureSheetWithHeaders_(ss, SHEET_BASE, BASE_HEADERS);
    const itensSh = ensureSheetWithHeaders_(ss, SHEET_ITENS, ITENS_HEADERS);

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
    const sh = ensureSheetWithHeaders_(ss, SHEET_LOG, LOG_HEADERS);
    sh.appendRow([new Date(), userCode, userEmail, acao, idBase, seqItem || '', detalhes || '']);
  }

  // ------------- API pública -------------
  return {
    getParam,
    validarPIN,
    handleUpload,
    listar,
    detalhar,
    dashboard,
    dashboardPessoal,
    minhasAtividades,
    validar,
    logAction,
    parseXml: parseXml_,
    saveRecord: saveRecord_
  };
})();
