/* Código do Google Apps Script que o gerador e a cadência usam (copiado pelo
   usuário para o script.google.com dele). Fica num arquivo só para as duas
   páginas mostrarem sempre a mesma versão. */
window.TAKE3_SCRIPT_VERSAO = 2;
window.TAKE3_SCRIPT_GMAIL = String.raw`// Take3 · envio direto e cadência pelo seu Gmail (Google Apps Script) · versão 2
// Troque a senha abaixo por uma só sua e use a mesma no gerador.
var SENHA = 'troque-esta-senha';
var VERSAO = 2;
var CAB = ['id', 'lead', 'email', 'nome', 'empresa', 'etapa', 'assunto', 'corpo', 'quando', 'status', 'enviadoEm', 'erro', 'criadoEm'];

// Rode esta uma vez (selecione autorizar e clique em Executar): dá a permissão,
// cria a planilha da fila e liga o relógio que envia os agendados a cada 5 minutos.
function autorizar() {
  var ss = planilha();
  garantirGatilho();
  GmailApp.getInboxUnreadCount();
  Logger.log('Tudo certo. Hoje ainda dá para enviar ' + MailApp.getRemainingDailyQuota() + ' e-mails. Planilha da fila: ' + ss.getUrl());
}

// Abrir a URL do app no navegador mostra esta mensagem: serve para conferir.
function doGet() {
  return ContentService.createTextOutput('Envio direto Take3: funcionando (versão 2). Agora cole esta URL no gerador.');
}

function doPost(e) {
  var r;
  try {
    var d = JSON.parse(e.postData.contents);
    if (SENHA === 'troque-esta-senha' || d.senha !== SENHA) throw new Error('Senha errada.');
    var acao = d.acao || (d.teste ? 'teste' : 'enviar');
    if (acao === 'teste') r = { ok: true };
    else if (acao === 'enviar') { enviar(d.para, d.assunto, d.corpo, d.remetente); r = { ok: true }; }
    else if (acao === 'agendar') r = { ok: true, agendados: agendar(d.itens || [], d.remetente) };
    else if (acao === 'listar') r = { ok: true, itens: listar(), planilha: planilha().getUrl() };
    else if (acao === 'parar') r = { ok: true, parados: pararLead(String(d.lead || '')) };
    else throw new Error('Ação desconhecida: atualize o script.');
    r.restante = MailApp.getRemainingDailyQuota();
    r.versao = VERSAO;
  } catch (err) {
    r = { ok: false, erro: String(err && err.message || err), versao: VERSAO };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function emailValido(x) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x || ''); }

function enviar(para, assunto, corpo, remetente) {
  if (!emailValido(para)) throw new Error('E-mail inválido.');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Limite diário do Gmail atingido. Continue amanhã.');
  var nome = String(remetente || PropertiesService.getScriptProperties().getProperty('remetente') || 'Take3 Digital').slice(0, 80);
  GmailApp.sendEmail(para, String(assunto || '').slice(0, 200), String(corpo || ''), { name: nome });
}

function planilha() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('planilha'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Take3 · Fila de e-mails');
    var aba = ss.getSheets()[0];
    aba.setName('fila');
    aba.appendRow(CAB);
    aba.setFrozenRows(1);
    props.setProperty('planilha', ss.getId());
  }
  return ss;
}

function abaFila() { var ss = planilha(); return ss.getSheetByName('fila') || ss.getSheets()[0]; }

function garantirGatilho() {
  var tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'processarFila'; });
  if (!tem) ScriptApp.newTrigger('processarFila').timeBased().everyMinutes(5).create();
}

function agendar(itens, remetente) {
  if (remetente) PropertiesService.getScriptProperties().setProperty('remetente', String(remetente).slice(0, 80));
  var criado = new Date(), linhas = [];
  itens.slice(0, 3000).forEach(function (x) {
    var quando = new Date(x.quando);
    if (!emailValido(x.email) || isNaN(quando.getTime())) return;
    linhas.push([Utilities.getUuid(), String(x.lead || x.email), String(x.email), String(x.nome || ''), String(x.empresa || ''),
      Number(x.etapa) || 1, String(x.assunto || '').slice(0, 200), String(x.corpo || ''), quando, 'AGENDADO', '', '', criado]);
  });
  if (linhas.length) {
    var aba = abaFila();
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, CAB.length).setValues(linhas);
  }
  garantirGatilho();
  return linhas.length;
}

function lerFila() {
  var aba = abaFila(), n = aba.getLastRow();
  return { aba: aba, linhas: n < 2 ? [] : aba.getRange(2, 1, n - 1, CAB.length).getValues() };
}

function data(v) { return v instanceof Date ? v : new Date(v); }
function iso(v) { if (v === '' || v === null || v === undefined) return ''; var d = data(v); return isNaN(d.getTime()) ? '' : d.toISOString(); }

function listar() {
  return lerFila().linhas.slice(-2000).map(function (l) {
    return { id: l[0], lead: l[1], email: l[2], nome: l[3], empresa: l[4], etapa: l[5], assunto: l[6],
      quando: iso(l[8]), status: l[9], enviadoEm: iso(l[10]), erro: l[11] };
  });
}

function marcarLead(fila, lead, status) {
  var n = 0;
  fila.linhas.forEach(function (l, i) {
    if (l[1] === lead && l[9] === 'AGENDADO') { l[9] = status; fila.aba.getRange(i + 2, 10).setValue(status); n++; }
  });
  return n;
}

function pararLead(lead) { return marcarLead(lerFila(), lead, 'CANCELADO'); }

// Respondeu depois que a sequência começou? E-mail voltou (endereço não existe)?
function verificarLead(email, desde) {
  var s = Math.floor((isNaN(desde.getTime()) ? Date.now() - 30 * 864e5 : desde.getTime()) / 1000);
  if (GmailApp.search('from:' + email + ' after:' + s, 0, 1).length) return 'RESPONDEU';
  if (GmailApp.search('from:mailer-daemon "' + email + '" after:' + s, 0, 1).length) return 'INVALIDO';
  return '';
}

// Roda sozinho a cada 5 minutos (o relógio criado em autorizar).
function processarFila() {
  var trava = LockService.getScriptLock();
  if (!trava.tryLock(1000)) return;
  try {
    var fila = lerFila(), agora = new Date(), enviados = 0, checado = {};
    for (var i = 0; i < fila.linhas.length && enviados < 25; i++) {
      var l = fila.linhas[i];
      if (l[9] !== 'AGENDADO' || data(l[8]) > agora) continue;
      var lead = l[1];
      if (checado[lead] === undefined) checado[lead] = verificarLead(l[2], data(l[12]));
      if (checado[lead]) { marcarLead(fila, lead, checado[lead]); continue; }
      if (MailApp.getRemainingDailyQuota() < 1) break;
      try {
        enviar(l[2], l[6], l[7], null);
        l[9] = 'ENVIADO';
        fila.aba.getRange(i + 2, 10, 1, 2).setValues([['ENVIADO', new Date()]]);
        enviados++;
        Utilities.sleep(1500);
      } catch (err) {
        l[9] = 'FALHOU';
        fila.aba.getRange(i + 2, 10, 1, 3).setValues([['FALHOU', new Date(), String(err && err.message || err)]]);
      }
    }
  } finally {
    trava.releaseLock();
  }
}
`;
