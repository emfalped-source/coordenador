/**
 * Coordenador Virtual — servidor (Google Apps Script)
 * Uso pessoal do coordenador. Os dados ficam numa planilha no Drive da conta que publica este script.
 *
 * Instalação (uma vez):
 * 1. Implantar > Nova implantação > App da Web: Executar como "Eu"; Quem pode acessar "Qualquer pessoa".
 *    Copie o endereço que termina em /exec.
 * 2. No primeiro acesso pelo app, a senha digitada vira a senha do app (guardada só como hash).
 *    A planilha "Coordenador Virtual — dados" é criada sozinha no seu Drive.
 *    Para trocar a senha depois: troque SENHA_INICIAL e rode "definirSenha".
 */
var SENHA_INICIAL = 'troque-esta-senha';
var NOME_PLANILHA = 'Coordenador Virtual — dados';
var DIAS_SESSAO = 365;
// Planilha de referência dos contatos (COORDENAÇÃO 2026 RELAÇÃO ALUNOS/RESPONSÁVEIS).
// O app lê os estudantes e telefones direto dela; basta manter essa planilha atualizada.
var PLANILHA_REFERENCIA = '1BmHXhXhWODvIrsstNwgxRm6YBBRDjkqztzqw8wLjfJs';
var REFERENCIA_MINUTOS = 60; // de quanto em quanto tempo o app relê a planilha de referência

// ---------- instalação ----------
function instalar() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('PLANILHA');
  if (id) { try { SpreadsheetApp.openById(id); Logger.log('Já instalado: ' + SpreadsheetApp.openById(id).getUrl()); return; } catch (e) {} }
  var ss = SpreadsheetApp.create(NOME_PLANILHA);
  var c = ss.getSheets()[0]; c.setName('Contatos');
  c.getRange(1, 1, 1, 6).setValues([['id', 'nome', 'turma', 'responsavel', 'contatos', 'tratamento']]);
  var k = ss.insertSheet('Casos');
  k.getRange(1, 1, 1, 13).setValues([['id', 'data', 'estudante_id', 'estudante', 'turma', 'tipo', 'situacao', 'nivel', 'canal', 'decisao', 'detalhes', 'mensagem', 'registrado_em']]);
  [c, k].forEach(function (s) { s.setFrozenRows(1); s.getRange('1:1').setFontWeight('bold'); });
  p.setProperty('PLANILHA', ss.getId());
  Logger.log('Planilha criada: ' + ss.getUrl());
}

function definirSenha() {
  if (!SENHA_INICIAL || SENHA_INICIAL === 'troque-esta-senha') throw new Error('Troque o texto de SENHA_INICIAL antes de rodar.');
  var sal = Utilities.getUuid();
  var p = PropertiesService.getScriptProperties();
  p.setProperty('SAL', sal);
  p.setProperty('SENHA', hash_(SENHA_INICIAL, sal));
  p.deleteProperty('SESSOES');
  Logger.log('Senha definida. Agora apague a senha da linha SENHA_INICIAL e salve.');
}

// ---------- web ----------
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = API[req.fn];
    if (!fn) throw new Error('Função desconhecida.');
    out = { ok: true, res: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, erro: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('Coordenador Virtual: servidor no ar.'); }

var API = {
  entrar: function (senha) {
    var p = PropertiesService.getScriptProperties();
    Utilities.sleep(400); // freia tentativas em série
    if (!p.getProperty('PLANILHA')) instalar();
    if (!p.getProperty('SENHA')) {
      // primeiro acesso: a senha digitada passa a ser a senha do app
      if (String(senha || '').length < 6) throw new Error('Primeiro acesso: crie uma senha com pelo menos 6 caracteres.');
      var sal = Utilities.getUuid();
      p.setProperty('SAL', sal);
      p.setProperty('SENHA', hash_(String(senha), sal));
    }
    if (hash_(String(senha || ''), p.getProperty('SAL')) !== p.getProperty('SENHA')) throw new Error('Senha incorreta.');
    var tk = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    var s = sessoes_(); s[hash_(tk, 'tk')] = Date.now() + DIAS_SESSAO * 864e5;
    gravarSessoes_(s);
    return { token: tk };
  },
  sair: function (tk) { var s = sessoes_(); delete s[hash_(String(tk), 'tk')]; gravarSessoes_(s); return true; },
  carregar: function (tk) {
    exigir_(tk);
    var info = null;
    var ult = Number(PropertiesService.getScriptProperties().getProperty('REF_ULTIMA') || 0);
    if (PLANILHA_REFERENCIA && Date.now() - ult > REFERENCIA_MINUTOS * 60000) { try { info = sincronizarReferencia_(); } catch (e) { info = { erro: String(e.message || e) }; } }
    return { alunos: lerAlunos_(), casos: lerCasos_(), referencia: info || statusReferencia_() };
  },
  atualizarReferencia: function (tk) {
    exigir_(tk);
    return { info: sincronizarReferencia_(), alunos: lerAlunos_() };
  },
  salvarAlunos: function (tk, lista) {
    exigir_(tk);
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = aba_('Contatos');
      var trat = {}; lerAlunos_().forEach(function (a) { if (a.g) trat[a.id] = a.g; });
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
      var rows = (lista || []).map(function (a) { return linhaAluno_(a, a.g || trat[a.id] || ''); });
      if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
      return rows.length;
    } finally { lock.releaseLock(); }
  },
  salvarAluno: function (tk, a) {
    exigir_(tk);
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = aba_('Contatos'), r = acharLinha_(sh, a.id), row = [linhaAluno_(a, a.g || '')];
      if (r) sh.getRange(r, 1, 1, 6).setValues(row); else sh.appendRow(row[0]);
      return true;
    } finally { lock.releaseLock(); }
  },
  salvarCaso: function (tk, c) {
    exigir_(tk);
    var sh = aba_('Casos');
    sh.appendRow([c.id, c.data, c.aid, c.nome, c.turma, c.cat, c.sub, c.nivel, c.canal, c.dec, c.det || '', c.msg || '', new Date()]);
    return true;
  },
  apagarCaso: function (tk, id) {
    exigir_(tk);
    var sh = aba_('Casos'), r = acharLinha_(sh, id);
    if (r) sh.deleteRow(r);
    return true;
  }
};

// ---------- dados ----------
function planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty('PLANILHA');
  if (!id) throw new Error('Servidor não instalado. Rode "instalar" no Apps Script.');
  return SpreadsheetApp.openById(id);
}
function aba_(n) { return planilha_().getSheetByName(n); }
function linhaAluno_(a, g) {
  return [String(a.id), String(a.nome || ''), String(a.turma || ''), String(a.resp || ''), JSON.stringify(a.cont || []), g || ''];
}
function lerAlunos_() {
  var sh = aba_('Contatos'), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    var cont = []; try { cont = JSON.parse(r[4] || '[]'); } catch (e) {}
    var a = { id: String(r[0]), nome: r[1], turma: r[2], resp: r[3], cont: cont, tels: cont.map(function (c) { return c.tel; }) };
    if (r[5]) a.g = r[5];
    return a;
  });
}
function lerCasos_() {
  var sh = aba_('Casos'), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var ini = Math.max(2, sh.getLastRow() - 999);
  return sh.getRange(ini, 1, sh.getLastRow() - ini + 1, 12).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    return { id: String(r[0]), data: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), aid: String(r[2]), nome: r[3], turma: r[4], cat: r[5], sub: r[6], nivel: Number(r[7]), canal: r[8], dec: r[9], det: r[10], msg: r[11] };
  }).reverse();
}
function acharLinha_(sh, id) {
  var n = sh.getLastRow() - 1; if (n < 1) return 0;
  var f = sh.getRange(2, 1, n, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

// ---------- segurança ----------
function hash_(txt, sal) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + txt, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function sessoes_() {
  var s = {}; try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty('SESSOES') || '{}'); } catch (e) {}
  var agora = Date.now(); Object.keys(s).forEach(function (k) { if (s[k] < agora) delete s[k]; });
  return s;
}
function gravarSessoes_(s) { PropertiesService.getScriptProperties().setProperty('SESSOES', JSON.stringify(s)); }
function exigir_(tk) {
  var s = sessoes_();
  if (!tk || !s[hash_(String(tk), 'tk')]) throw new Error('SESSAO: Sessão expirada. Entre de novo.');
}

// ---------- planilha de referência ----------
function statusReferencia_() {
  var p = PropertiesService.getScriptProperties();
  return { quando: Number(p.getProperty('REF_ULTIMA') || 0), total: Number(p.getProperty('REF_TOTAL') || 0) };
}
function norm_(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
function tc_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().replace(/^[-\s]+|[-\s]+$/g, '').toLowerCase().split(' ')
    .map(function (w, i) { return i && ['de', 'da', 'do', 'das', 'dos', 'e'].indexOf(w) >= 0 ? w : w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
}
function mkId_(n) { return norm_(n).replace(/[^a-z0-9]+/g, '-').slice(0, 60); }
function normTels_(v) {
  var out = [];
  String(v == null ? '' : v).split(/[,\/;?]| e | ou /i).forEach(function (p) {
    var d = p.replace(/\D/g, '').replace(/^0+/, '');
    if (d.indexOf('55') === 0 && d.length >= 12) d = d.slice(2);
    if (d.length === 8 || d.length === 9) d = '31' + d;
    if (d.length === 10 && /[6-9]/.test(d[2])) d = d.slice(0, 2) + '9' + d.slice(2);
    var f = '(' + d.slice(0, 2) + ') ' + d.slice(2, -4) + '-' + d.slice(-4);
    if (d.length === 11 && d[2] === '9') out.push({ tel: f, wa: true });
    else if (d.length === 10 && /[2-5]/.test(d[2])) out.push({ tel: f, wa: false });
  });
  return out;
}
function sincronizarReferencia_() {
  var ss = SpreadsheetApp.openById(PLANILHA_REFERENCIA);
  var novos = [], vistos = {};
  ss.getSheets().forEach(function (sh) {
    var sn = sh.getName();
    var rows = sh.getDataRange().getValues().map(function (r) { return r.map(function (c) { return typeof c === 'number' ? String(Math.round(c)) : String(c == null ? '' : c); }); });
    var hi = -1;
    for (var k = 0; k < rows.length && k < 15; k++) { if (rows[k].some(function (c) { return /^(aluno|estudante|nome)/.test(norm_(c)); })) { hi = k; break; } }
    if (hi < 0) return;
    var h = rows[hi].map(norm_);
    var iN = -1, iT = -1, pares = [];
    h.forEach(function (x, i) {
      if (iN < 0 && /^(aluno|estudante|nome)/.test(x)) iN = i;
      if (iT < 0 && /turma|classe/.test(x)) iT = i;
      var m = x.match(/^(mae|pai|respons)/); if (m) pares.push([i, { mae: 'Mãe', pai: 'Pai', respons: 'Responsável' }[m[1]]]);
    });
    rows.slice(hi + 1).forEach(function (r) {
      var nome = tc_(r[iN]);
      var turma = iT >= 0 ? String(r[iT]).trim().toUpperCase() : '';
      if (!turma && nome.split(' ').length >= 2) turma = String(sn).trim().toUpperCase();
      if (!nome || !turma || /^\d+$/.test(nome)) return;
      var cont = [];
      pares.forEach(function (pq) {
        var pn = r[pq[0]] && r[pq[0]] !== '0' ? tc_(r[pq[0]]) : ''; if (pn === '-') pn = '';
        normTels_(r[pq[0] + 1]).forEach(function (t) { cont.push({ q: pq[1], nome: pn, tel: t.tel, wa: t.wa }); });
      });
      var id = mkId_(nome); if (vistos[id]) id += '-' + mkId_(turma); vistos[id] = 1;
      var c0 = cont.filter(function (c) { return c.nome; })[0];
      novos.push({ id: id, nome: nome, turma: turma, resp: c0 ? c0.q + ': ' + c0.nome : '', cont: cont });
    });
  });
  if (!novos.length) throw new Error('A planilha de referência não tem estudantes legíveis.');
  novos.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt'); });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh2 = aba_('Contatos');
    var trat = {}; lerAlunos_().forEach(function (a) { if (a.g) trat[a.id] = a.g; });
    if (sh2.getLastRow() > 1) sh2.getRange(2, 1, sh2.getLastRow() - 1, 6).clearContent();
    var linhas = novos.map(function (a) { return linhaAluno_(a, trat[a.id] || ''); });
    sh2.getRange(2, 1, linhas.length, 6).setValues(linhas);
  } finally { lock.releaseLock(); }
  var p = PropertiesService.getScriptProperties();
  p.setProperty('REF_ULTIMA', String(Date.now())); p.setProperty('REF_TOTAL', String(novos.length));
  return { quando: Date.now(), total: novos.length, comWhats: novos.filter(function (a) { return a.cont.some(function (c) { return c.wa; }); }).length };
}
// Rode no editor para testar a leitura da planilha de referência.
function testarReferencia() { Logger.log(JSON.stringify(sincronizarReferencia_())); }
