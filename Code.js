/*
 * Código principal do projeto GAS_DevolucoesNFe (versão PIN por planilha Users).
 * Roteia requisições, trata CORS (opcional), autentica via PIN fixo (aba Users)
 * e delega a lógica ao Services.gs.
 */

// Nomes de abas/tabelas
const SHEET_BASE   = 'Base';
const SHEET_ITENS  = 'Itens';
const SHEET_LOG    = 'Log';
const SHEET_PARAMS = 'Params';
const SHEET_USERS  = 'Users';

const PROJECT_NAME = 'GAS_DevolucoesNFe';

/** GET */
function doGet(e) {

  var action = (e && e.parameter && e.parameter.action) || '';
  
  // Se não houver action, entrega a interface (Frontend.html)
  if (!action) {
    return HtmlService.createTemplateFromFile('Frontend')
      .evaluate()
      .setTitle('Devoluções NFe')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Caso contrário, trata como chamada de API JSON
  return handleRequest(e, false);

}

/** POST */
function doPost(e) {
  return handleRequest(e, true);
}


// chama a validação de PIN do Services no lado do servidor
function apiLoginServer(pin) {
  try {
    if (!pin) return { ok:false, message:'PIN vazio' };
    if (!Services || typeof Services.validarPIN !== 'function') {
      return { ok:false, message:'Services.validarPIN não encontrado' };
    }
    return Services.validarPIN(pin); // { ok:true, pin, nome } ou { ok:false, message }
  } catch (err) {
    return { ok:false, message:String(err && err.message || err) };
  }
}

/**
 * Router + (opcional) CORS + autenticação por PIN fixo (aba Users).
 */
function handleRequest(e, isPost) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';
    const origin = (e && e.headers && e.headers.origin) ? e.headers.origin : '*';

    // Lê CORS permitido dos Params (ou usa '*')
    const corsAllowed = (typeof Services.getParam === 'function'
      ? (Services.getParam('CORS_ALLOWED_ORIGINS') || '*')
      : '*');

    const response = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);

    const sendJson = function (obj) {
      const jsonStr = JSON.stringify(obj);
      response.setContent(jsonStr);

      // --- CORS opcional (remova se não precisar / se gerar erro no seu deploy) ---
      try {
        const allowedOrigins = (corsAllowed || '*').split(',').map(function (s) { return s.trim(); });
        if (corsAllowed === '*' || allowedOrigins.indexOf(origin) !== -1) {
          // Algumas execuções do GAS Web App não suportam set de headers arbitrários.
          // Se lançar erro aqui, comente estas 3 linhas abaixo.
          response.addHeader && response.addHeader('Access-Control-Allow-Origin', origin);
          response.addHeader && response.addHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
          response.addHeader && response.addHeader('Access-Control-Allow-Headers', 'Content-Type');
        }
      } catch (err) {
        // silenciosamente ignora problema de header em ambientes onde não é suportado
      }
      // ------------------------------------------------------------------------------

      return response;
    };

    // Pré-flight CORS
    if (e && e.method && String(e.method).toLowerCase() === 'options') {
      return sendJson({ ok: true, message: 'CORS preflight' });
    }

    // --- Rotas públicas (sem exigir PIN) ---
    if (action === 'login' || action === 'verificarPIN') {
      const pinPublic = params.pin || '';
      if (!pinPublic) return sendJson({ ok: false, code: 'PIN_REQUIRED', message: 'Informe o PIN.' });

      // Valida na aba Users
      if (typeof Services.validarPIN !== 'function') {
        return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.validarPIN não encontrado.' });
      }
      const result = Services.validarPIN(pinPublic); // { ok, nome?, ... }
      return sendJson(result);
    }

    // --- Demais rotas exigem PIN válido ---
    let pin = '';
    if (params.pin) {
      pin = params.pin;
    } else if (isPost && e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        pin = body.pin || '';
      } catch (err) {
        // segue vazio
      }
    }
    if (!pin) {
      return sendJson({ ok: false, code: 'PIN_REQUIRED', message: 'PIN é obrigatório.' });
    }

    if (typeof Services.validarPIN !== 'function') {
      return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.validarPIN não encontrado.' });
    }
    const session = Services.validarPIN(pin); // { ok:true, nome:'...' } quando válido
    if (!session || !session.ok) {
      return sendJson({ ok: false, code: 'SESSION_INVALID', message: 'PIN inválido ou inativo.' });
    }

    // Identidade do usuário na sessão
    const userCode  = pin;
    const userName  = session.nome || '';
    const userEmail = (Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';

    // Router de ações
    switch (action) {
      case 'uploadXml':
        if (typeof Services.handleUpload !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.handleUpload não encontrado.' });
        return sendJson(Services.handleUpload(e, userCode, userEmail, userName));

      case 'listar':
        if (typeof Services.listar !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.listar não encontrado.' });
        return sendJson(Services.listar(params, userCode, userName));

      case 'detalhar':
        if (typeof Services.detalhar !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.detalhar não encontrado.' });
        return sendJson(Services.detalhar(params.id || '', userCode, userName));

      case 'validar': {
        if (typeof Services.validar !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.validar não encontrado.' });
        let payload = {};
        if (isPost && e.postData && e.postData.contents) {
          try {
            payload = JSON.parse(e.postData.contents);
          } catch (err) {
            return sendJson({ ok: false, code: 'BAD_JSON', message: 'Payload inválido.' });
          }
        }
        return sendJson(Services.validar(payload, userCode, userEmail, userName));
      }

      case 'salvarAnexo':
        if (typeof Services.salvarAnexo !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.salvarAnexo não encontrado.' });
        return sendJson(Services.salvarAnexo(e, userCode, userName));

      case 'dashboard':
        if (typeof Services.dashboard !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.dashboard não encontrado.' });
        return sendJson(Services.dashboard(params));

      case 'dashboardPessoal':
        if (typeof Services.dashboardPessoal !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.dashboardPessoal não encontrado.' });
        return sendJson(Services.dashboardPessoal(userCode, userName));

      case 'minhasAtividades':
        if (typeof Services.minhasAtividades !== 'function') return sendJson({ ok: false, code: 'MISSING_SERVICE', message: 'Services.minhasAtividades não encontrado.' });
        return sendJson(Services.minhasAtividades(params, userCode, userName));

      default:
        return sendJson({ ok: false, code: 'ACTION_UNKNOWN', message: 'Ação desconhecida.' });
    }

  } catch (err) {
    const response = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
    response.setContent(JSON.stringify({
      ok: false,
      code: 'EXCEPTION',
      message: err && err.message ? err.message : String(err),
      stack: err && err.stack ? err.stack : ''
    }));
    return response;
  }
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === Helpers de teste, chamados via google.script.run ===
function apiUploadServer(xml, pin) {
  try {
    if (!pin || !xml) return { ok:false, message:'PIN e XML são obrigatórios' };
    // Monta um "e" compatível com Services.handleUpload (JSON com campo "xml")
    const e = { postData: { type: 'application/json', contents: JSON.stringify({ xml: xml }) } };
    return Services.handleUpload(e, pin, '');
  } catch (err) {
    return { ok:false, message:String(err && err.message || err) };
  }
}

const WEBAPP_VERSION = 'v0.3-listar-fix';

function apiVersionServer() {
  return { ok: true, version: WEBAPP_VERSION, ts: new Date().toISOString() };
}

function apiListarServer(pin) {
  try {
    if (!pin) return { ok:false, code:'PIN_REQUIRED', message:'PIN obrigatório' };
    const out = Services.listar({}, pin);
    if (!out || out.ok !== true) {
      return {
        ok:false,
        code: (out && out.code) || 'LISTAR_FAIL',
        message: (out && out.message) || 'Falha em Services.listar'
      };
    }
    return out; // <- IMPORTANTE: sempre retornar
  } catch (e) {
    Logger.log('apiListarServer error: ' + (e && e.stack || e));
    return { ok:false, code:'EXCEPTION', message: String(e && e.message || e) };
  }
}

function apiDetalharServer(id, pin) {
  try {
    Logger.log('[apiDetalharServer] id=%s pin=%s', id, pin);

    if (!pin || !id) {
      return { ok:false, code:'MISSING_ARGS', message:'PIN e ID obrigatórios' };
    }
    if (typeof Services === 'undefined') {
      return { ok:false, code:'NO_SERVICES', message:'Objeto Services não encontrado' };
    }
    if (typeof Services.detalhar !== 'function') {
      return { ok:false, code:'MISSING_METHOD', message:'Services.detalhar não é uma função' };
    }

    var out = Services.detalhar(id, pin);
    Logger.log('[apiDetalharServer] retorno=%s', JSON.stringify(out));

    if (out == null) { // pega null/undefined
      return { ok:false, code:'NULL_RETURN', message:'Services.detalhar retornou null/undefined' };
    }
    return out;

  } catch (e) {
    Logger.log('[apiDetalharServer][EXCEPTION] %s\n%s', e && e.message, e && e.stack);
    return {
      ok:false,
      code:'EXCEPTION',
      message: String(e && e.message || e),
      stack: e && e.stack ? String(e.stack) : ''
    };
  }
}



function apiValidarServer(payload, pin) {
  try { if (!pin) return { ok:false, message:'PIN obrigatório' };
    return Services.validar(payload, pin, '', payload.validadorNome || '');
  } catch (e) { return { ok:false, message:String(e) }; }
}

function apiDashboardServer(pin) {
  try {
    if (!pin) return { ok:false, code:'PIN_REQUIRED', message:'PIN obrigatório' };
    const session = Services.validarPIN(pin);
    if (!session || !session.ok) {
      return { ok:false, code:'SESSION_INVALID', message:'PIN inválido ou inativo.' };
    }
    const resp = Services.dashboard({}) || {};
    return resp.ok ? resp : Object.assign({ ok:false }, resp, { message: resp.message || 'Falha ao montar dashboard.' });
  } catch (e) {
    Logger.log('[apiDashboardServer][EXCEPTION] %s', e && e.stack || e);
    return { ok:false, code:'EXCEPTION', message:String(e && e.message || e) };
  }
}

function apiDashboardPessoalServer(pin) {
  try {
    if (!pin) return { ok:false, code:'PIN_REQUIRED', message:'PIN obrigatório' };
    const session = Services.validarPIN(pin);
    if (!session || !session.ok) {
      return { ok:false, code:'SESSION_INVALID', message:'PIN inválido ou inativo.' };
    }
    const resp = Services.dashboardPessoal(pin, session.nome || '') || {};
    return resp.ok ? resp : Object.assign({ ok:false }, resp, { message: resp.message || 'Falha ao montar dashboard pessoal.' });
  } catch (e) {
    Logger.log('[apiDashboardPessoalServer][EXCEPTION] %s', e && e.stack || e);
    return { ok:false, code:'EXCEPTION', message:String(e && e.message || e) };
  }
}

function testListar() {
  var pin = "123456"; // coloca um PIN válido que exista na aba Users
  var res = apiListarServer(pin);
  Logger.log(JSON.stringify(res, null, 2));
}

function testServicesListar() {
  const pin = "123456";
  const out = Services.listar({}, pin);
  Logger.log(JSON.stringify(out,null,2));
}

function apiListarPing() {
  return { ok:true, rows:[{ ChaveNFe:'TESTE', Emissao:'2025-01-01', Emitente_Nome:'Demo', ValorNF:'10.00', Status:'Pendente' }], total:1 };
}

function apiPingServer() {
  // deve NUNCA ser null
  return { ok: true, pong: new Date().toISOString() };
}

function apiDetalharServer2(id, pin) {
  try {
    Logger.log('[apiDetalharServer2] id=%s pin=%s', id, pin);

    if (!pin || !id) {
      return { ok:false, code:'MISSING_ARGS', message:'PIN e ID obrigatórios' };
    }
    if (typeof Services === 'undefined') {
      return { ok:false, code:'NO_SERVICES', message:'Objeto Services não encontrado' };
    }
    if (typeof Services.detalhar !== 'function') {
      return { ok:false, code:'MISSING_METHOD', message:'Services.detalhar não é uma função' };
    }

    var out = Services.detalhar(id, pin);
    Logger.log('[apiDetalharServer2] retorno=%s', JSON.stringify(out));

    if (out == null) {
      // <- aqui evitamos "null" no cliente
      return { ok:false, code:'NULL_RETURN', message:'Services.detalhar retornou null/undefined' };
    }
    return out;

  } catch (e) {
    Logger.log('[apiDetalharServer2][EXCEPTION] %s\n%s', e && e.message, e && e.stack);
    return {
      ok:false,
      code:'EXCEPTION',
      message: String(e && e.message || e),
      stack: e && e.stack ? String(e.stack) : ''
    };
  }
}


