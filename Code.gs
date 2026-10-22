/**
 * Combo Zero – backend do app na Planilha Google
 *
 * 1. Crie uma Planilha Google (ex.: "Combo Zero – Respostas").
 * 2. Extensões > Apps Script > cole este código > salve.
 * 3. Implantar > Nova implantação > Tipo: App da Web
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa  (o app do vendedor não faz login)
 * 4. Copie a URL que termina em /exec e cole em window.SHEETS_URL no app.html.
 *
 * Aba "Respostas": uma linha por loja, sempre com a situação mais recente.
 * Aba "Histórico": toda marcação, com data e hora, para auditoria.
 */

const ABA = 'Respostas';
const HIST = 'Histórico';
const COLS = ['Cod', 'Cliente', 'Gerência', 'Supervisor', 'Rota', 'Situação no piloto',
              'Kit no pedido', 'Combo no ar', 'Fora do iFood', 'Observação', 'Data da observação', 'Última marcação'];

function aba_(nome) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.appendRow(nome === HIST ? ['Registrado em'].concat(COLS) : COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  }
  return sh;
}

/** O app lê aqui a situação de todas as lojas. */
function doGet() {
  const sh = aba_(ABA);
  const v = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    out[String(r[0])] = {
      g: r[2], kit: r[6] === true, publicado: r[7] === true, saiu: r[8] === true,
      obs: String(r[9] || ''), obsAt: r[10] ? new Date(r[10]).toISOString() : '',
      at: r[11] ? new Date(r[11]).toISOString() : ''
    };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** O app grava aqui cada marcação. */
function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  if (!d.c) return ContentService.createTextOutput('sem código');
  const sit = { acao: 'Piloto com combo', controle: 'Comparação do piloto', aberto: 'Aberta' }[d.st] || '';
  const row = [String(d.c), d.n || '', d.g || '', d.sv || '', d.ro || '', sit,
               !!d.kit, !!d.publicado, !!d.saiu, String(d.obs || '').slice(0, 500),
               d.obsAt ? new Date(d.obsAt) : '', d.at ? new Date(d.at) : new Date()];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba_(ABA);
    const n = sh.getLastRow();
    const cods = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
    const i = cods.indexOf(String(d.c));
    if (i < 0) sh.appendRow(row); else sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    aba_(HIST).appendRow([new Date()].concat(row));
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
