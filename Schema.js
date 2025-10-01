/*
 * Funções auxiliares de criação e manutenção do schema da planilha.
 * Inclui menu personalizado para criação da estrutura e acesso ao
 * frontend diretamente pela planilha (útil em ambiente de testes).
 */

/**
 * Função onOpen adiciona menu customizado à planilha para que o
 * administrador possa executar a configuração inicial (setup) e
 * abrir a interface web a partir da planilha. Esta função é
 * executada automaticamente quando a planilha é aberta.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Devolucoes NFe')
    .addItem('Configurar estrutura', 'setup')
    .addItem('Abrir Interface', 'showFrontend')
    .addToUi();
}

/**
 * Cria a planilha, abas e parâmetros se ainda não existirem. Esta
 * função pode ser executada manualmente pelo administrador caso
 * a estrutura seja alterada. Utiliza a função getSpreadsheet() em
 * Services para garantir consistência.
 */
function setup() {
  // Cria (ou abre) a planilha principal
  var ss = SpreadsheetApp.create("DevolucoesNFe");
  
  // Aba Base
  var base = ss.insertSheet("Base");
  base.appendRow([
    "ID","ChaveNFe","Numero","Serie","Emissao",
    "Emitente_Nome","Emitente_CNPJ","Destinatario_Nome","Destinatario_CNPJ",
    "CFOP","ValorNF","ValorICMS",
    "XML_DriveURL","Status",
    "CriadoPorCode","CriadoEm",
    "ValidadoPorCode","ValidadorNome","ValidadoEm",
    "FormaPagamento","DataPagamento",
    "Anexo_DriveURL","Observacoes","AtualizadoEm"
  ]);

  // Aba Itens
  var itens = ss.insertSheet("Itens");
  itens.appendRow([
    "ID_RegistroBase","Seq","Codigo","Descricao","NCM","CFOP",
    "Quantidade","VlrUnit","VlrTotal",
    "StatusItem","MotivoItem","ObsItem","QtdeDevolvida",
    "ValidadoPorCode_Item","ValidadoEm_Item"
  ]);

  // Aba Log
  var log = ss.insertSheet("Log");
  log.appendRow([
    "Timestamp","UserCode","UsuarioEmail","Acao","ID_RegistroBase","SeqItem","Detalhes"
  ]);

  // Aba Params
  var params = ss.insertSheet("Params");
  params.appendRow(["Param","Valor"]);
  params.appendRow(["PIN_TTL_MINUTES","15"]);
  params.appendRow(["CACHE_KPI_MIN","5"]);
  params.appendRow(["MAX_UPLOAD_MB","20"]);
  params.appendRow(["ALLOW_ZIP","true"]);
  params.appendRow(["EXIGIR_ANEXO_AO_ACEITAR","false"]);
  params.appendRow(["CORS_ALLOWED_ORIGINS","*"]);

  // Remove a aba padrão "Sheet1" se existir
  var sheet = ss.getSheetByName("Sheet1");
  if (sheet) ss.deleteSheet(sheet);

  Logger.log("Planilha criada com ID: " + ss.getId());
}
function ensureSchema() {
  var ss = SpreadsheetApp.openById(PLANILHA_ID);

  // --- Users ---
  var users = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  if (users.getLastRow() === 0) {
    users.appendRow(['PIN','NomeUsuario','Ativo']);
    users.appendRow(['123456','Usuário Demo', true]);
  }

  // --- Params ---
  var params = ss.getSheetByName(SHEET_PARAMS) || ss.insertSheet(SHEET_PARAMS);
  if (params.getLastRow() === 0) {
    params.appendRow(['Param','Valor']);
    params.appendRow(['CORS_ALLOWED_ORIGINS','*']);
  }

  // --- Base ---
  var base = ss.getSheetByName(SHEET_BASE) || ss.insertSheet(SHEET_BASE);
  if (base.getLastRow() === 0) base.appendRow([
    'ID','ChaveNFe','Numero','Serie','Emissao',
    'Emitente_Nome','Emitente_CNPJ','Destinatario_Nome','Destinatario_CNPJ',
    'CFOP','ValorNF','ValorICMS','XML_DriveURL','Status',
    'CriadoPorCode','CriadoEm','ValidadoPorCode','ValidadorNome','ValidadoEm',
    'FormaPagamento','DataPagamento','Anexo_DriveURL','Observacoes','AtualizadoEm'
  ]);

  // --- Itens ---
  var itens = ss.getSheetByName(SHEET_ITENS) || ss.insertSheet(SHEET_ITENS);
  if (itens.getLastRow() === 0) itens.appendRow([
    'ID_RegistroBase','Seq','Codigo','Descricao','NCM','CFOP','Quantidade','VlrUnit','VlrTotal',
    'StatusItem','MotivoItem','ObsItem','QtdeDevolvida','ValidadoPorCode_Item','ValidadoEm_Item'
  ]);

  // --- Log ---
  var log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) log.appendRow([
    'Timestamp','UserCode','UsuarioEmail','Acao','ID_RegistroBase','SeqItem','Detalhes'
  ]);
}


/**
 * Abre a interface frontend (HTML) dentro da planilha. Útil para
 * testes locais. Essa função usa o HtmlService para servir o
 * conteúdo de Frontend.html, aplicando também o CSS e JS.
 */
function showFrontend() {
  const html = HtmlService.createHtmlOutputFromFile('Frontend')
    .setTitle('Devoluções NFe')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showSidebar(html);
}