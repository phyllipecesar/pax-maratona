// Pax · inscrições da Maratona
// Cole em Extensões → Apps Script da planilha de finalistas, salve e publique como
// App da Web (executar como você, acesso: qualquer pessoa). Depois de salvar, recarregue a
// planilha: aparece o menu "Pax" com "Gerar IDs dos maratonistas". Rode uma vez.
//
// O que este script faz:
//   gerarIds()        cria a coluna "uuid" na aba de maratonistas e preenche os vazios.
//   doGet(?roster=1)  devolve a lista de maratonistas com uuid (o formulário carrega daqui).
//   doPost            grava uma resposta na aba "Respostas", com o uuid (ou MANUAL_ADDED).

const ABA_RESPOSTAS = 'Respostas';
const ABA_ROSTER = ''; // vazio = primeira aba da planilha. Ou o nome exato da aba de maratonistas.
const COL_ID = 'uuid';
const COLS = ['id', 'ts', 'nome', 'faculdade', 'sigla', 'time', 'posicao', 'na_lista', 'telefone', 'email', 'conhecia', 'interesses', 'achou', 'quem', 'origem'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pax')
    .addItem('Gerar IDs dos maratonistas', 'gerarIds')
    .addToUi();
}

function rosterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ABA_ROSTER ? ss.getSheetByName(ABA_ROSTER) : ss.getSheets()[0];
}

function headerIndex_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const idx = {};
  h.forEach((name, i) => { if (name) idx[name] = i; });
  return { names: h, idx: idx };
}

// Cria a coluna uuid (no fim) se não existir e preenche as linhas que têm Nome mas não têm uuid.
function gerarIds() {
  const sh = rosterSheet_();
  let { names, idx } = headerIndex_(sh);
  if (idx[COL_ID] === undefined) {
    sh.getRange(1, names.length + 1).setValue(COL_ID);
    ({ names, idx } = headerIndex_(sh));
  }
  const colNome = idx['Nome'];
  if (colNome === undefined) throw new Error('Não achei a coluna "Nome" na aba ' + sh.getName());
  const last = sh.getLastRow();
  if (last < 2) return;
  const nomes = sh.getRange(2, colNome + 1, last - 1, 1).getValues();
  const ids = sh.getRange(2, idx[COL_ID] + 1, last - 1, 1).getValues();
  let n = 0;
  for (let i = 0; i < ids.length; i++) {
    if (!ids[i][0] && String(nomes[i][0]).trim()) { ids[i][0] = Utilities.getUuid(); n++; }
  }
  sh.getRange(2, idx[COL_ID] + 1, last - 1, 1).setValues(ids);
  try { SpreadsheetApp.getUi().alert('IDs gerados: ' + n + ' (total de linhas: ' + ids.length + ')'); } catch (e) {}
  return n;
}

// Lista de maratonistas com uuid, para o autocomplete do formulário.
function roster_() {
  const sh = rosterSheet_();
  const { idx } = headerIndex_(sh);
  const need = ['Nome', 'Faculdade', 'Sigla', 'Time', 'Posição 1ª fase', COL_ID];
  for (const k of need) if (idx[k] === undefined) throw new Error('Coluna ausente: ' + k + ' (rode "Gerar IDs" primeiro)');
  const last = sh.getLastRow();
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (const r of rows) {
    const nome = String(r[idx['Nome']]).trim();
    const id = String(r[idx[COL_ID]]).trim();
    if (!nome || !id || nome === '(nome não encontrado)') continue;
    out.push({ id: id, n: nome, f: String(r[idx['Faculdade']]), u: String(r[idx['Sigla']]), t: String(r[idx['Time']]), p: Number(r[idx['Posição 1ª fase']]) || null });
  }
  return out;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.roster) {
    return ContentService.createTextOutput(JSON.stringify(roster_())).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('ok');
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_RESPOSTAS);
  if (!sh) { sh = ss.insertSheet(ABA_RESPOSTAS); sh.appendRow(COLS); sh.setFrozenRows(1); }
  // Garante o cabeçalho atual (se a aba foi criada por uma versão antiga, reescreve a linha 1).
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), COLS.length)).getValues()[0];
  if (String(h[0]) !== COLS[0]) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { d = (e && e.parameter) || {}; }
  if (!d.id) d.id = d.na_lista ? 'SEM_ID' : 'MANUAL_ADDED';
  const row = COLS.map(k => Array.isArray(d[k]) ? d[k].join(', ') : (d[k] === undefined || d[k] === null ? '' : d[k]));
  sh.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: d.id })).setMimeType(ContentService.MimeType.JSON);
}
