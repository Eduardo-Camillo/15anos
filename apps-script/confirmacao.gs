/**
 * Recebe as confirmações de presença do convite e salva numa planilha do Google.
 *
 * COMO CONFIGURAR
 * 1. Crie uma planilha nova no Google Sheets (ex.: "Confirmações - 15 anos Milena").
 * 2. Na planilha: Extensões > Apps Script.
 * 3. Apague o que estiver lá, cole este arquivo inteiro e salve (ícone de disquete).
 * 4. Implantar > Nova implantação > engrenagem > App da Web
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 *    Clique em Implantar e autorize o acesso à sua conta quando o Google pedir.
 * 5. Copie a URL gerada (termina em /exec) e cole em URL_CONFIRMACAO no index.html.
 *
 * As abas "Confirmações" e "Resumo" são criadas sozinhas na primeira resposta.
 * Se a mesma pessoa (mesmo nome) responder de novo, a linha dela é ATUALIZADA,
 * em vez de duplicada — assim a contagem de convidados não fica errada.
 */

// Proteção simples contra envios indevidos. Use o MESMO texto no index.html.
const CHAVE_EVENTO = 'Milena15';

const ABA = 'Confirmações';
const CABECALHO = ['Data/hora', 'Nome', 'Vai à festa?', 'Acompanhantes', 'Nomes dos acompanhantes', 'Total de pessoas'];
const MAX_ACOMPANHANTES = 10;

function doPost(e) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);   // evita que duas respostas simultâneas se atropelem
  try {
    const dados = JSON.parse(e.postData.contents);
    if (dados.chave !== CHAVE_EVENTO) return resposta({ ok: false, erro: 'chave inválida' });

    const nome = limpar(dados.nome);
    if (!nome) return resposta({ ok: false, erro: 'nome vazio' });

    const vai = dados.vai === 'sim';
    const acompanhantes = vai && Array.isArray(dados.acompanhantes)
      ? dados.acompanhantes.map(limpar).filter(String).slice(0, MAX_ACOMPANHANTES)
      : [];

    const linha = [
      new Date(),
      nome,
      vai ? 'Sim' : 'Não',
      acompanhantes.length,
      acompanhantes.join(', '),
      vai ? 1 + acompanhantes.length : 0
    ];

    const aba = obterAba();
    const existente = procurarNome(aba, nome);
    if (existente) aba.getRange(existente, 1, 1, linha.length).setValues([linha]);
    else aba.appendRow(linha);

    return resposta({ ok: true, atualizado: !!existente });
  } catch (erro) {
    return resposta({ ok: false, erro: String(erro) });
  } finally {
    trava.releaseLock();
  }
}

// Abrir a URL no navegador só para testar se está no ar.
function doGet() {
  return resposta({ ok: true, mensagem: 'Confirmação de presença funcionando.' });
}

function obterAba() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA, 0);
    aba.appendRow(CABECALHO);
    aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#E8EDE9');
    aba.setFrozenRows(1);
    aba.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    aba.setColumnWidth(2, 220);
    aba.setColumnWidth(5, 320);
    criarResumo(planilha);
  }
  return aba;
}

function criarResumo(planilha) {
  if (planilha.getSheetByName('Resumo')) return;
  const r = planilha.insertSheet('Resumo', 1);
  const a = "'" + ABA + "'!";
  r.getRange('A1:B5').setValues([
    ['Total de pessoas confirmadas', ''],
    ['Respostas "Sim"', ''],
    ['Respostas "Não"', ''],
    ['Acompanhantes', ''],
    ['Total de respostas', '']
  ]);
  r.getRange('B1').setFormula('=SUM(' + a + 'F2:F)');
  // SUMPRODUCT com um único argumento funciona em qualquer idioma da planilha
  // (em português o separador é ";", em inglês é "," — COUNTIF quebraria num dos dois)
  r.getRange('B2').setFormula('=SUMPRODUCT(' + a + 'C2:C="Sim")');
  r.getRange('B3').setFormula('=SUMPRODUCT(' + a + 'C2:C="Não")');
  r.getRange('B4').setFormula('=SUM(' + a + 'D2:D)');
  r.getRange('B5').setFormula('=COUNTA(' + a + 'B2:B)');
  r.getRange('A1:B1').setFontWeight('bold');
  r.setColumnWidth(1, 240);
  // remove a "Página1" vazia que vem com toda planilha nova
  const vazia = planilha.getSheetByName('Página1') || planilha.getSheetByName('Sheet1');
  if (vazia && vazia.getLastRow() === 0) planilha.deleteSheet(vazia);
}

// Procura a pessoa pelo nome (sem diferenciar maiúsculas/acentos). Retorna o nº da linha ou 0.
function procurarNome(aba, nome) {
  const ultima = aba.getLastRow();
  if (ultima < 2) return 0;
  const alvo = normalizar(nome);
  const nomes = aba.getRange(2, 2, ultima - 1, 1).getValues();
  for (let i = 0; i < nomes.length; i++) {
    if (normalizar(nomes[i][0]) === alvo) return i + 2;
  }
  return 0;
}

function normalizar(texto) {
  return String(texto || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function limpar(texto) {
  let t = String(texto || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  // impede que um nome começando com = + - @ vire fórmula na planilha
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function resposta(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
