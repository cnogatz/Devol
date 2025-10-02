/*
 * Coleção de testes simples para validação das funções principais.
 * Estes testes podem ser executados manualmente pelo desenvolvedor
 * para verificar o correto funcionamento da leitura de XML, salvamento
 * de registros e listagem. Execute runTests() no editor do Apps
 * Script para verificar os resultados no log.
 */
//teste
function runTests() {
  testParseXml();
  testSaveRecord();
  testListar();
}

function testParseXml() {
  const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<nfeProc>\n  <NFe>\n    <infNFe Id="NFe123" versao="4.00">\n      <ide>\n        <nNF>1</nNF>\n        <serie>1</serie>\n        <dhEmi>2025-01-01T08:30:00-03:00</dhEmi>\n      </ide>\n      <emit>\n        <CNPJ>12345678000195</CNPJ>\n        <xNome>Fornecedor Teste</xNome>\n      </emit>\n      <dest>\n        <CNPJ>98765432000100</CNPJ>\n        <xNome>Loja Teste</xNome>\n      </dest>\n      <det nItem="1">\n        <prod>\n          <cProd>XYZ</cProd>\n          <xProd>Produto Teste</xProd>\n          <NCM>1234567</NCM>\n          <CFOP>5949</CFOP>\n          <qCom>1</qCom>\n          <vUnCom>5.00</vUnCom>\n          <vProd>5.00</vProd>\n        </prod>\n      </det>\n      <total>\n        <ICMSTot>\n          <vNF>5.00</vNF>\n          <vICMS>0.00</vICMS>\n        </ICMSTot>\n      </total>\n    </infNFe>\n  </NFe>\n</nfeProc>`;
  const result = Services.parseXml(xml);
  Logger.log('Teste parseXml - base: %s', JSON.stringify(result.base));
  Logger.log('Itens: %s', JSON.stringify(result.itens));
}

function testSaveRecord() {
  const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<nfeProc><NFe><infNFe Id="NFeABC" versao="4.00"><ide><nNF>2</nNF><serie>1</serie><dhEmi>2025-02-01T09:00:00-03:00</dhEmi></ide><emit><CNPJ>11111111000100</CNPJ><xNome>Fornecedor 2</xNome></emit><dest><CNPJ>22222222000100</CNPJ><xNome>Destinatário 2</xNome></dest><det nItem="1"><prod><cProd>CD1</cProd><xProd>Item CD1</xProd><NCM>9876543</NCM><CFOP>5929</CFOP><qCom>3</qCom><vUnCom>2.00</vUnCom><vProd>6.00</vProd></prod></det><total><ICMSTot><vNF>6.00</vNF><vICMS>0.00</vICMS></ICMSTot></total></infNFe></NFe></nfeProc>`;
  const parsed = Services.parseXml(xml);
  const id = Services.saveRecord(parsed.base, parsed.itens, '999999');
  Logger.log('Teste saveRecord - novo ID: %s', id);
}

function testListar() {
  const result = Services.listar({ status: '' }, '999999');
  Logger.log('Teste listar - total %s, linhas %s', result.total, result.rows.length);
}
