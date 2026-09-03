/**
 * ================================================
 *  GAME LIBRARY MANAGER — by TM's Enterprise
 *  Backend (Google Apps Script)
 * ================================================
 */

const SHEET_NAME = 'JOGOS';
const HEADERS = ['ID', 'Nome', 'Biblioteca', 'Valor', 'Zerado', 'Platinado', 'Tempo Jogado', 'Data de Aquisição', 'Observações', 'Pretende Platinar'];

const BIBLIOTECAS = [
  'Steam', 'Epic Games', 'GOG', 'Amazon Games', 'Ubisoft Connect',
  'EA App', 'Battle.net', 'Microsoft Store', 'PlayStation', 'Xbox', 'Nintendo', 'Outro'
];

/**
 * Ponto de entrada do Web App.
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle("Hadess's Game Library Manager - by TM's Enterprise")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Permite que index.html inclua style.html e script.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtém (ou cria) a aba JOGOS, garantindo o cabeçalho.
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Planilha já existente: só garante que a coluna "Pretende Platinar" exista no cabeçalho,
    // sem tocar em nenhuma linha ou coluna já cadastrada.
    const ultimoCabecalho = sheet.getRange(1, HEADERS.length).getValue();
    if (!ultimoCabecalho) {
      sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]);
    }
  }
  return sheet;
}

/**
 * Converte uma linha da planilha (array) em objeto.
 */
function rowToObject_(row) {
  let tempoJogado = '';
  if (row[6] instanceof Date) {
    // O Google Sheets converteu automaticamente um valor como "7.2" em uma data (dia.mês).
    // Reconstrói o texto original a partir da data para não perder a informação.
    tempoJogado = row[6].getDate() + '.' + (row[6].getMonth() + 1);
  } else if (row[6]) {
    tempoJogado = String(row[6]);
  }

  return {
    id: String(row[0] || ''),
    nome: String(row[1] || ''),
    biblioteca: String(row[2] || ''),
    valor: Number(row[3]) || 0,
    zerado: (row[4] === 'Sim') ? 'Sim' : 'Não',
    platinado: (row[5] === 'Sim') ? 'Sim' : 'Não',
    tempoJogado: tempoJogado,
    dataAquisicao: row[7] instanceof Date ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(row[7] || ''),
    observacoes: row[8] ? String(row[8]) : '',
    pretendePlatinar: (row[9] === 'Sim') ? 'Sim' : 'Não'
  };
}

/**
 * Gera o próximo ID sequencial (ex: "001", "002"...).
 */
function gerarProximoId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const nums = ids
    .map(id => parseInt(String(id).replace(/\D/g, ''), 10))
    .filter(n => !isNaN(n));
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return String(next).padStart(3, '0');
}

/**
 * Lista todos os jogos cadastrados.
 */
function listarJogos() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return data
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(rowToObject_);
}

/**
 * RN001 / RN002 / RN005 / RN006 — validação de campos obrigatórios.
 */
function validarJogo_(dados) {
  if (!dados.nome || !String(dados.nome).trim()) {
    throw new Error('O nome do jogo é obrigatório.');
  }
  if (!dados.biblioteca || !String(dados.biblioteca).trim()) {
    throw new Error('A biblioteca é obrigatória.');
  }
  dados.zerado = (dados.zerado === 'Sim') ? 'Sim' : 'Não';
  dados.platinado = (dados.platinado === 'Sim') ? 'Sim' : 'Não';
  dados.pretendePlatinar = (dados.pretendePlatinar === true || dados.pretendePlatinar === 'Sim') ? 'Sim' : 'Não';
  return dados;
}

/**
 * Adiciona um novo jogo à planilha.
 */
function adicionarJogo(dados) {
  dados = validarJogo_(dados);
  const sheet = getSheet_();
  const id = gerarProximoId_(sheet);
  const valor = (dados.valor === '' || dados.valor === undefined || dados.valor === null) ? 0 : Number(dados.valor);

  sheet.appendRow([
    id,
    String(dados.nome).trim(),
    dados.biblioteca,
    valor,
    dados.zerado,
    dados.platinado,
    dados.tempoJogado || '',
    new Date(),
    dados.observacoes || '',
    dados.pretendePlatinar
  ]);

  return { sucesso: true, id: id, mensagem: 'Jogo adicionado com sucesso!' };
}

/**
 * Encontra a linha (1-based) de um jogo pelo ID.
 */
function encontrarLinhaPorId_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => String(v) === String(id));
  return idx === -1 ? -1 : idx + 2;
}

/**
 * Edita um jogo existente.
 */
function editarJogo(id, dados) {
  dados = validarJogo_(dados);
  const sheet = getSheet_();
  const linha = encontrarLinhaPorId_(sheet, id);
  if (linha === -1) throw new Error('Jogo não encontrado.');

  const valor = (dados.valor === '' || dados.valor === undefined || dados.valor === null) ? 0 : Number(dados.valor);
  const atual = sheet.getRange(linha, 8, 1, 1).getValue(); // preserva data de aquisição original

  sheet.getRange(linha, 2, 1, HEADERS.length - 1).setValues([[
    String(dados.nome).trim(),
    dados.biblioteca,
    valor,
    dados.zerado,
    dados.platinado,
    dados.tempoJogado || '',
    atual || new Date(),
    dados.observacoes || '',
    dados.pretendePlatinar
  ]]);

  return { sucesso: true, mensagem: 'Jogo atualizado com sucesso!' };
}

/**
 * Exclui um jogo pelo ID.
 */
function excluirJogo(id) {
  const sheet = getSheet_();
  const linha = encontrarLinhaPorId_(sheet, id);
  if (linha === -1) throw new Error('Jogo não encontrado.');
  sheet.deleteRow(linha);
  return { sucesso: true, mensagem: 'Jogo excluído com sucesso!' };
}

/**
 * RN007 — Pesquisa parcial e case-insensitive por nome.
 */
function buscarJogos(termo) {
  const jogos = listarJogos();
  if (!termo || !termo.trim()) return jogos;
  const termoNormalizado = termo.trim().toLowerCase();
  return jogos.filter(j => String(j.nome).toLowerCase().indexOf(termoNormalizado) !== -1);
}

/**
 * Retorna a lista de bibliotecas suportadas (para o <select>).
 */
function listarBibliotecas() {
  return BIBLIOTECAS;
}

/**
 * Converte strings como "82h", "125h 30min", "82 horas" em horas totais (float).
 */
function parseHoras_(tempo) {
  if (!tempo) return 0;
  const str = String(tempo).toLowerCase();
  let horas = 0;
  let encontrou = false;

  const matchH = str.match(/(\d+(?:[.,]\d+)?)\s*h(?!oras)/);
  const matchHoras = str.match(/(\d+(?:[.,]\d+)?)\s*horas?/);
  const matchMin = str.match(/(\d+)\s*min/);

  if (matchH) { horas += parseFloat(matchH[1].replace(',', '.')); encontrou = true; }
  else if (matchHoras) { horas += parseFloat(matchHoras[1].replace(',', '.')); encontrou = true; }

  if (matchMin) { horas += parseInt(matchMin[1], 10) / 60; encontrou = true; }

  if (!encontrou) {
    const soNumero = parseFloat(str.replace(',', '.'));
    if (!isNaN(soNumero)) horas = soNumero;
  }
  return horas;
}

/**
 * Estatísticas para o dashboard.
 */
function getDashboardStats() {
  const jogos = listarJogos();

  const totalJogos = jogos.length;
  const totalZerados = jogos.filter(j => j.zerado === 'Sim').length;
  const totalPlatinados = jogos.filter(j => j.platinado === 'Sim').length;
  const totalInvestido = jogos.reduce((acc, j) => acc + (Number(j.valor) || 0), 0);
  const totalHoras = jogos.reduce((acc, j) => acc + parseHoras_(j.tempoJogado), 0);

  const porBiblioteca = {};
  jogos.forEach(j => {
    const bib = j.biblioteca || 'Outro';
    if (!porBiblioteca[bib]) {
      porBiblioteca[bib] = { quantidade: 0, valorInvestido: 0 };
    }
    porBiblioteca[bib].quantidade++;
    porBiblioteca[bib].valorInvestido += Number(j.valor) || 0;
  });

  return {
    totalJogos,
    totalZerados,
    totalPlatinados,
    totalInvestido,
    totalHoras: Math.round(totalHoras * 10) / 10,
    porBiblioteca
  };
}
