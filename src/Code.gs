/**
 * PARTE DIÁRIA DE EQUIPAMENTOS
 * App de apontamento de equipamentos (Google Apps Script + Planilha Google).
 *
 * Como instalar: veja o passo a passo enviado junto. Resumo:
 *  1) Planilha Google nova > Extensões > Apps Script.
 *  2) Cole este arquivo em Code.gs e crie o arquivo HTML "Index" com o outro código.
 *  3) Rode a função  configurar  uma vez (autorize).
 *  4) Implantar > Nova implantação > App da Web (Executar como: eu · Acesso: Qualquer pessoa).
 */

const TZ = 'America/Sao_Paulo';

const CONFIG_PADRAO = {
  nome: 'Minha Obra',
  intervaloMin: 60,
  emailChefe: 'gerente@exemplo.com',
  nomeChefe: 'Gerente',
  emailEngenheira: 'engenharia@exemplo.com',
  assinatura: 'Eng. Responsável',
  pinApontador: '1234',   // troque na aba Config depois de instalar
  pinAdmin: '0000'        // troque na aba Config depois de instalar
};

const COL_EQ = ['id', 'prefixo', 'descricao', 'medidor', 'operador', 'empresa', 'valorMensal', 'inicio', 'saida'];
const COL_PT = ['id', 'data', 'equipId', 'prefixo', 'descricao', 'operador', 'empresa', 'medidor', 'horIni', 'horFim',
                'hIni', 'hFim', 'abastLitros', 'paradas', 'obs', 'criadoEm', 'atualizadoEm'];

/* ======================= INSTALAÇÃO ======================= */

function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  ss.setSpreadsheetTimeZone(TZ);

  const eq = aba_(ss, 'Equipamentos', COL_EQ);
  const pt = aba_(ss, 'Partes', COL_PT);
  const cf = aba_(ss, 'Config', ['chave', 'valor']);

  if (cf.getLastRow() < 2) {
    const linhas = Object.keys(CONFIG_PADRAO).map(k => [k, String(CONFIG_PADRAO[k])]);
    cf.getRange(2, 1, linhas.length, 2).setValues(linhas);
  }
  if (eq.getLastRow() < 2) {
    const linhas = EQUIP_INICIAL.map(e => COL_EQ.map(c => e[c] == null ? '' : String(e[c])));
    eq.getRange(2, 1, linhas.length, COL_EQ.length).setValues(linhas);
  }
  if (pt.getLastRow() < 2 && PARTE_INICIAL) {
    const p = PARTE_INICIAL;
    pt.getRange(2, 1, 1, COL_PT.length).setValues([COL_PT.map(c => c === 'paradas' ? JSON.stringify(p.paradas) : (p[c] == null ? '' : String(p[c])))]);
  }
  const s1 = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 3) ss.deleteSheet(s1);

  agendarEnvio();

  Logger.log('Pronto! Agora faça: Implantar > Nova implantação > App da Web.');
  Logger.log('PIN do apontador: ' + cfg_().pinApontador + '  ·  PIN de administradora: ' + cfg_().pinAdmin + '  (troque na aba Config)');
}

/** Agenda o e-mail diário para as 08h00 (apaga o agendamento antigo). Rode uma vez depois de colar o código novo. */
function agendarEnvio() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'enviarDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarDiario').timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone(TZ).create();
  Logger.log('E-mail diário agendado para as 08h00 (segunda a sexta; na segunda vai a parte de sexta).');
}

function aba_(ss, nome, cab) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  sh.getRange(1, 1, sh.getMaxRows(), Math.max(cab.length, 2)).setNumberFormat('@');
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#E3EFEE');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ======================= WEB APP ======================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Parte Diária')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function cfg_() {
  const sh = ss_().getSheetByName('Config');
  const out = Object.assign({}, CONFIG_PADRAO);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim(); });
  out.intervaloMin = Number(out.intervaloMin) || 0;
  return out;
}

function nivel_(pin) {
  const c = cfg_();
  pin = String(pin || '').trim();
  if (pin && pin === String(c.pinAdmin)) return 'admin';
  if (pin && pin === String(c.pinApontador)) return 'apontador';
  throw new Error('PIN');
}

function ler_(nome, cols) {
  const sh = ss_().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; cols.forEach((c, i) => o[c] = r[i] === '' ? null : String(r[i])); return o; });
}

const NUM_PT = ['horIni', 'horFim', 'abastLitros'];
const NUM_EQ = ['valorMensal'];
function normParte_(p) {
  NUM_PT.forEach(k => { p[k] = p[k] == null ? null : Number(p[k]); });
  try { p.paradas = p.paradas ? JSON.parse(p.paradas) : []; } catch (e) { p.paradas = []; }
  return p;
}
function equipamentos_() { return ler_('Equipamentos', COL_EQ).map(e => { NUM_EQ.forEach(k => e[k] = e[k] == null ? null : Number(e[k])); return e; }); }
function partes_(de, ate) { return ler_('Partes', COL_PT).map(normParte_).filter(p => (!de || p.data >= de) && (!ate || p.data <= ate)); }

function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function diasAtras_(n) { return Utilities.formatDate(new Date(Date.now() - n * 864e5), TZ, 'yyyy-MM-dd'); }

/* ---------- chamadas do app ---------- */

function apiInit(pin) {
  const nv = nivel_(pin);
  const c = cfg_();
  return {
    nivel: nv,
    cfg: nv === 'admin'
      ? { nome: c.nome, intervaloMin: c.intervaloMin, emailEngenheira: c.emailEngenheira, emailChefe: c.emailChefe, nomeChefe: c.nomeChefe }
      : { nome: c.nome, intervaloMin: c.intervaloMin },
    // valor mensal dos equipamentos só vai para a administradora
    equip: nv === 'admin' ? equipamentos_() : equipamentos_().map(e => Object.assign({}, e, { valorMensal: null })),
    partes: partes_(diasAtras_(62), null),
    planilha: nv === 'admin' ? ss_().getUrl() : ''
  };
}

function apiPartes(pin, de, ate) { nivel_(pin); return partes_(de, ate); }

function apiUltimoHorimetro(pin, equipId, antesDe) {
  nivel_(pin);
  const l = partes_(null, null).filter(p => p.equipId === equipId && p.data < antesDe && p.horFim != null).sort((a, b) => b.data.localeCompare(a.data));
  return l[0] || null;
}

function apiSalvarParte(pin, p) {
  nivel_(pin);
  if (!p || !/^\d{4}-\d{2}-\d{2}$/.test(p.data || '') || !p.equipId) throw new Error('Dados incompletos');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName('Partes');
    const id = p.data + '_' + p.equipId;
    const agora = new Date().toISOString();
    const achou = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    const linhaAnt = achou ? sh.getRange(achou.getRow(), 1, 1, COL_PT.length).getValues()[0] : null;
    const doc = Object.assign({}, p, {
      id: id,
      paradas: JSON.stringify(p.paradas || []),
      criadoEm: linhaAnt ? linhaAnt[COL_PT.indexOf('criadoEm')] : agora,
      atualizadoEm: agora
    });
    const valores = [COL_PT.map(c => doc[c] == null ? '' : String(doc[c]))];
    if (achou) sh.getRange(achou.getRow(), 1, 1, COL_PT.length).setValues(valores);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, COL_PT.length).setNumberFormat('@').setValues(valores);
    return id;
  } finally { lock.releaseLock(); }
}

function apiExcluirParte(pin, id) {
  nivel_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName('Partes');
    const achou = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (achou) sh.deleteRow(achou.getRow());
    return true;
  } finally { lock.releaseLock(); }
}

function apiSalvarEquip(pin, e) {
  if (nivel_(pin) !== 'admin') throw new Error('Só a administradora pode alterar o cadastro.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName('Equipamentos');
    if (!e.id) e.id = String(e.prefixo).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') + '-' + String(e.inicio).replace(/-/g, '');
    const v = [COL_EQ.map(c => e[c] == null ? '' : String(e[c]))];
    const achou = sh.getRange('A:A').createTextFinder(e.id).matchEntireCell(true).findNext();
    if (achou) sh.getRange(achou.getRow(), 1, 1, COL_EQ.length).setValues(v);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, COL_EQ.length).setNumberFormat('@').setValues(v);
    return equipamentos_();
  } finally { lock.releaseLock(); }
}

function apiSalvarCfg(pin, c) {
  if (nivel_(pin) !== 'admin') throw new Error('Só a administradora pode alterar os parâmetros.');
  const sh = ss_().getSheetByName('Config');
  const dados = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  Object.keys(c).forEach(k => {
    const i = dados.findIndex(r => r[0] === k);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(String(c[k]));
    else sh.appendRow([k, String(c[k])]);
  });
  return true;
}

/** Envia PDF + Excel BDE do período escolhido para o e-mail da engenheira. */
function apiEnviarRelatorio(pin, de, ate, empresa, equipId) {
  if (nivel_(pin) !== 'admin') throw new Error('Só a administradora pode enviar relatórios.');
  const c = cfg_();
  const ps = partes_(de, ate).filter(p => (!empresa || p.empresa === empresa) && (!equipId || p.equipId === equipId));
  if (!ps.length) throw new Error('Não há partes no período.');
  const per = br_(de) + ' a ' + br_(ate);
  const pdf = gerarPdf_(ps, 'PARTE DIÁRIA DE EQUIPAMENTOS – CONSOLIDADO', c.nome + '   |   Período: ' + per + (empresa ? '   |   Empresa: ' + empresa : ''))
    .setName('Parte_Diaria_' + de + '_a_' + ate + '.pdf');
  const xls = gerarBde_(ps, de, ate, c).setName('BDE_' + de + '_a_' + ate + '.xlsx');
  MailApp.sendEmail({ to: c.emailEngenheira, subject: 'Relatório Parte Diária – ' + c.nome + ' – ' + per, name: 'Parte Diária',
    body: 'Relatório do período ' + per + ' (' + ps.length + ' partes). PDF para assinatura e BDE com HT/HD/HP em anexo.', attachments: [pdf, xls] });
  return c.emailEngenheira;
}

/* ======================= ENVIO DIÁRIO (gatilho 08:00) ======================= */

/** Dia da parte a enviar: terça a sexta = dia anterior; segunda = sexta; sábado e domingo = não envia. */
function diaDoEnvio_() {
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u')); // 1 = segunda ... 7 = domingo
  if (dow === 6 || dow === 7) return null;
  return diasAtras_(dow === 1 ? 3 : 1);
}

function enviarDiario() {
  const dia = diaDoEnvio_();
  if (!dia) return; // fim de semana: sem envio
  const c = cfg_();
  const mesIni = dia.slice(0, 8) + '01';
  const doMes = partes_(mesIni, dia);
  const doDia = doMes.filter(p => p.data === dia);
  const d = br_(dia);
  const xls = gerarBde_(doMes, mesIni, dia, c).setName('BDE_' + dia.slice(0, 7) + '_ate_' + dia + '.xlsx');

  if (doDia.length) {
    const pdf = gerarPdfDiario_(doDia, dia).setName('Parte_Diaria_' + dia + '.pdf');
    MailApp.sendEmail({ to: c.emailChefe, name: c.assinatura,
      subject: 'Parte Diária – ' + c.nome + ' – ' + d,
      body: 'Bom dia, ' + c.nomeChefe + '.\n\nSegue em anexo a parte diária dos equipamentos da ' + c.nome + ' do dia ' + d + ' (' + doDia.length +
            ' equipamentos apontados), para impressão e assinatura: uma folha por equipamento e, no final, o resumo do dia.\n\nAtenciosamente,\n' + c.assinatura,
      attachments: [pdf] });
    MailApp.sendEmail({ to: c.emailEngenheira, name: 'Parte Diária',
      subject: 'Parte Diária + BDE – ' + c.nome + ' – ' + d,
      body: 'Parte diária de ' + d + ' (' + doDia.length + ' equipamentos) e o BDE do mês até essa data, com as linhas HT/HD/HP para copiar na planilha de apropriação.',
      attachments: [pdf.copyBlob().setName(pdf.getName()), xls] });
  } else {
    MailApp.sendEmail({ to: c.emailEngenheira, name: 'Parte Diária',
      subject: 'Parte Diária – ' + c.nome + ' – ' + d + ': nenhuma parte lançada',
      body: 'O apontador não lançou nenhuma parte em ' + d + ', por isso o e-mail para ' + c.nomeChefe + ' não foi enviado hoje. Segue o BDE do mês até essa data.',
      attachments: [xls] });
  }
}

/** Para testar o envio agora (manda só para a engenheira): parte diária do último dia com lançamentos. */
function testarEnvio() {
  const c = cfg_();
  const ps = partes_(diasAtras_(40), hoje_());
  if (!ps.length) { Logger.log('Sem partes nos últimos 40 dias.'); return; }
  const dia = ps.map(p => p.data).sort().pop();
  const doDia = ps.filter(p => p.data === dia);
  const pdf = gerarPdfDiario_(doDia, dia).setName('Teste_Parte_Diaria_' + dia + '.pdf');
  const xls = gerarBde_(ps.filter(p => p.data.slice(0, 7) === dia.slice(0, 7)), dia.slice(0, 8) + '01', dia, c).setName('Teste_BDE.xlsx');
  MailApp.sendEmail({ to: c.emailEngenheira, subject: '[TESTE] Parte Diária – ' + c.nome, body: 'Teste do envio pelo Google.', attachments: [pdf, xls] });
  Logger.log('Teste enviado para ' + c.emailEngenheira);
}

/* ======================= CÁLCULOS ======================= */

function mins_(t) { if (!t) return null; const m = String(t).match(/(\d{1,2}):(\d{2})/); return m ? Number(m[1]) * 60 + Number(m[2]) : null; }
function calc_(p, intervalo) {
  const a = mins_(p.hIni), b = mins_(p.hFim);
  const j = a != null && b != null ? b - a : null;
  const iv = j != null && j >= 360 ? intervalo : 0;
  let par = 0;
  (p.paradas || []).forEach(x => { const u = mins_(x.de), v = mins_(x.ate); if (u != null && v != null && v > u) par += v - u; });
  return { jornada: j, intervalo: iv, parada: par, efetivo: j != null ? Math.max(0, j - iv - par) : null, hd: j != null ? j - iv : null,
           dif: p.horIni != null && p.horFim != null ? p.horFim - p.horIni : null };
}
function hhmm_(m) { return m == null ? '—' : Math.floor(m / 60) + ':' + ('0' + (m % 60)).slice(-2); }
function br_(iso) { return iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4); }
function n1_(v) { return v == null || isNaN(v) ? '—' : Utilities.formatString('%.1f', v).replace('.', ','); }

/* ======================= ARQUIVOS (PDF e Excel) ======================= */

function tmp_(nome) { return SpreadsheetApp.create('tmp_' + nome + '_' + Date.now()); }
function exportar_(ss, fmt, extra) {
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=' + fmt + (extra || '');
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return blob;
}

const PDF_OPC = '&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=RIGHT&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
const LINHA_ASS = '________________________________________';

/** PDF consolidado (relatório por período): tabela única. */
function gerarPdf_(partes, titulo, subtitulo) {
  const c = cfg_();
  const ss = tmp_('pdf');
  montarTabela_(ss.getSheets()[0], partes, titulo, subtitulo, c);
  return exportar_(ss, 'pdf', PDF_OPC);
}

/** PDF do dia: uma folha por equipamento (para o operador assinar) e, no final, o resumo do dia. */
function gerarPdfDiario_(partes, dia) {
  const c = cfg_();
  const ss = tmp_('pdf');
  const ordem = partes.slice().sort((a, b) => (a.prefixo || '').localeCompare(b.prefixo || ''));
  ordem.forEach((p, i) => folhaParte_(i === 0 ? ss.getSheets()[0] : ss.insertSheet('P' + (i + 1), ss.getNumSheets()), p, c));
  const res = ss.insertSheet('Resumo', ss.getNumSheets());
  montarTabela_(res, ordem, 'RESUMO DO DIA – PARTE DIÁRIA DE EQUIPAMENTOS',
    c.nome + '   |   Data: ' + br_(dia) + '   |   ' + ordem.length + ' equipamento(s) apontado(s)', c);
  return exportar_(ss, 'pdf', PDF_OPC);
}

/** Uma folha com a parte diária de um equipamento. */
function folhaParte_(sh, p, c) {
  const W = 8;
  for (let i = 1; i <= W; i++) sh.setColumnWidth(i, 145);
  sh.getRange(1, 1, 60, W).setNumberFormat('@').setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1).setValue('PARTE DIÁRIA DE EQUIPAMENTO').setFontSize(14).setFontWeight('bold').setFontColor('#023334');
  sh.getRange(2, 1).setValue(c.nome + '   |   Data: ' + br_(p.data)).setFontColor('#444444');
  logo_(sh, 7);

  const k = calc_(p, c.intervaloMin);
  const u = p.medidor === 'km' ? 'km' : 'h';
  let r = 4;
  r = tabela_(sh, r, [['Prefixo', 1], ['Equipamento', 3], ['Operador', 2], ['Empresa', 2]],
    [[p.prefixo, p.descricao, p.operador, p.empresa]]);
  r = tabela_(sh, r, [['Inicial (' + u + ')', 1], ['Final (' + u + ')', 1], ['Diferença', 1], ['Abastecimento (L)', 1],
                      ['Início', 1], ['Término', 1], ['Jornada', 1], ['Intervalo', 1]],
    [[n1_(p.horIni), n1_(p.horFim), k.dif == null ? '—' : n1_(k.dif) + ' ' + u, n1_(p.abastLitros),
      p.hIni || '—', p.hFim || '—', hhmm_(k.jornada), k.intervalo ? hhmm_(k.intervalo) : '—']]);
  r = tabela_(sh, r, [['Horas disponíveis (HD)', 2], ['Horas paradas (HP)', 2], ['Horas efetivas (HT)', 2], ['Medidor', 2]],
    [[hhmm_(k.hd), k.parada ? hhmm_(k.parada) : '—', hhmm_(k.efetivo), u === 'km' ? 'Quilometragem (km)' : 'Horímetro (h)']]);

  const pars = (p.paradas || []).map(x => {
    const a = mins_(x.de), b = mins_(x.ate);
    return [x.de || '—', x.ate || '—', a != null && b != null && b > a ? hhmm_(b - a) : '—', x.motivo || ''];
  });
  r = tabela_(sh, r, [['Parada – de', 1], ['Até', 1], ['Duração', 1], ['Motivo da parada', 5]],
    pars.length ? pars : [['—', '—', '—', 'Sem paradas registradas']]);

  const obs = p.obs || '—';
  const rObs = r;
  r = tabela_(sh, r, [['Observações', 8]], [[obs]]);
  sh.setRowHeight(rObs + 1, Math.max(24, Math.ceil(obs.length / 140) * 18));

  const rs = r + 2;
  assinaturas_(sh, rs, [[1, 2, 'Apontador'], [3, 4, 'Operador' + (p.operador ? ' – ' + p.operador : '')],
                        [5, 6, 'Encarregado'], [7, 8, 'Gerente do contrato / Engenheiro']]);
  aparar_(sh, W, rs + 2);
}

/** Cabeçalho + linhas; cada coluna do cabeçalho é [rótulo, nº de colunas mescladas]. Devolve a próxima linha livre. */
function tabela_(sh, r, cab, linhas) {
  const W = cab.reduce((s, x) => s + x[1], 0);
  let col = 1;
  cab.forEach((x, i) => {
    cel_(sh, r, col, x[1], x[0]).setFontWeight('bold').setFontSize(8).setFontColor('#023334').setBackground('#E3EFEE');
    linhas.forEach((l, j) => cel_(sh, r + 1 + j, col, x[1], l[i]).setWrap(true));
    col += x[1];
  });
  sh.setRowHeights(r + 1, linhas.length, 24);
  sh.getRange(r, 1, linhas.length + 1, W).setBorder(true, true, true, true, true, true, '#9FB8B6', SpreadsheetApp.BorderStyle.SOLID);
  return r + linhas.length + 2;
}

function cel_(sh, r, c, n, v) {
  const g = sh.getRange(r, c, 1, n);
  if (n > 1) g.merge();
  g.setNumberFormat('@').setValue(v == null ? '' : String(v));
  return g;
}

function assinaturas_(sh, rs, grupos) {
  sh.setRowHeight(rs - 1, 40);
  grupos.forEach(g => {
    const n = g[1] - g[0] + 1;
    sh.getRange(rs, g[0], 1, n).merge().setValue(LINHA_ASS).setHorizontalAlignment('center');
    sh.getRange(rs + 1, g[0], 1, n).merge().setValue(g[2]).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(rs + 2, g[0], 1, n).merge().setValue('Data: ____/____/______').setHorizontalAlignment('center');
  });
}

function logo_(sh, col) {
  if (!LOGO_B64) return;
  try { sh.insertImage(Utilities.newBlob(Utilities.base64Decode(LOGO_B64), 'image/png', 'logo.png'), col, 1).setWidth(150).setHeight(40); } catch (e) {}
}

function aparar_(sh, W, ult) {
  if (sh.getMaxColumns() > W) sh.deleteColumns(W + 1, sh.getMaxColumns() - W);
  if (sh.getMaxRows() > ult) sh.deleteRows(ult + 1, sh.getMaxRows() - ult);
}

/** Tabela com todos os equipamentos (resumo do dia ou relatório do período). */
function montarTabela_(sh, partes, titulo, subtitulo, c) {
  const W = [62, 190, 105, 135, 72, 72, 70, 52, 58, 58, 56, 260];
  W.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, 1, 400, 12).setNumberFormat('@').setFontFamily('Arial').setFontSize(9).setVerticalAlignment('top');
  sh.getRange(1, 1).setValue(titulo).setFontSize(14).setFontWeight('bold').setFontColor('#023334');
  sh.getRange(2, 1).setValue(subtitulo).setFontColor('#444444');
  logo_(sh, 11);

  const cab = ['Prefixo', 'Equipamento', 'Operador', 'Empresa', 'Inicial', 'Final', 'Dif.', 'Início', 'Término', 'Efetivas', 'Paradas', 'Observações'];
  const ordem = partes.slice().sort((a, b) => (a.prefixo || '').localeCompare(b.prefixo || '') || a.data.localeCompare(b.data));
  const varios = new Set(ordem.map(p => p.data)).size > 1;
  if (varios) cab[0] = 'Data / Pref.';
  let te = 0, tp = 0;
  const linhas = ordem.map(p => {
    const k = calc_(p, c.intervaloMin); const u = p.medidor === 'km' ? 'km' : 'h';
    te += k.efetivo || 0; tp += k.parada;
    const par = (p.paradas || []).map(x => x.de + '–' + x.ate + ' ' + (x.motivo || '')).join('; ');
    return [(varios ? br_(p.data).slice(0, 5) + ' ' : '') + p.prefixo, p.descricao, p.operador, p.empresa, n1_(p.horIni), n1_(p.horFim),
            k.dif == null ? '—' : n1_(k.dif) + ' ' + u, p.hIni || '—', p.hFim || '—', hhmm_(k.efetivo), k.parada ? hhmm_(k.parada) : '—',
            [p.obs || '', par ? 'Paradas: ' + par : ''].filter(String).join(' | ')];
  });
  linhas.push(['Total', '', '', '', '', '', '', '', '', hhmm_(te), tp ? hhmm_(tp) : '—', '']);
  const r0 = 4;
  sh.getRange(r0, 1, 1, 12).setValues([cab]).setFontWeight('bold').setBackground('#E3EFEE');
  sh.getRange(r0 + 1, 1, linhas.length, 12).setValues(linhas).setWrap(true);
  sh.getRange(r0 + linhas.length, 1, 1, 12).setFontWeight('bold').setBackground('#F3F8F7');
  sh.getRange(r0, 5, linhas.length + 1, 7).setHorizontalAlignment('right');
  sh.getRange(r0, 1, linhas.length + 1, 12).setBorder(true, true, true, true, false, true, '#9FB8B6', SpreadsheetApp.BorderStyle.SOLID);

  const rs = r0 + linhas.length + 4;
  assinaturas_(sh, rs, [[1, 3, 'Apontador'], [5, 8, 'Encarregado'], [10, 12, 'Gerente do contrato / Engenheiro']]);
  aparar_(sh, 12, rs + 2);
}

function gerarBde_(partes, de, ate, c) {
  const ss = tmp_('bde'); const ws = ss.getSheets()[0]; ws.setName('BDE (copiar)');
  const d0 = new Date(Number(ate.slice(0, 4)), Number(ate.slice(5, 7)) - 1, 1);
  const ndias = new Date(d0.getFullYear(), d0.getMonth() + 1, 0).getDate();
  const L = [];
  const fmt = [];
  L.push(['BDE – HT / HD / HP por dia · ' + c.nome + ' · mês ' + ate.slice(5, 7) + '/' + ate.slice(0, 4) + ' (partes até ' + br_(ate) + ')']);
  L.push(['Copie as linhas HT e HD de cada equipamento para a sua planilha (horas [h]:mm). Dias em branco = sem parte lançada.']);
  L.push(['']);
  const grupos = {};
  partes.forEach(p => (grupos[p.equipId] = grupos[p.equipId] || []).push(p));
  Object.keys(grupos).sort((a, b) => (grupos[a][0].prefixo || '').localeCompare(grupos[b][0].prefixo || '')).forEach(id => {
    const ps = grupos[id], p0 = ps[0];
    L.push([p0.prefixo + ' – ' + p0.descricao + '  |  ' + p0.operador + '  |  ' + p0.empresa]); fmt.push({ r: L.length, t: 'grp' });
    const dias = ['DIA']; for (let d = 1; d <= ndias; d++) dias.push(d); dias.push('TOTAL'); L.push(dias); fmt.push({ r: L.length, t: 'hdr' });
    const por = {}; ps.forEach(p => { if (p.data.slice(0, 7) === ate.slice(0, 7)) por[Number(p.data.slice(8, 10))] = calc_(p, c.intervaloMin); });
    [['HT', 'efetivo'], ['HD', 'hd'], ['HP', 'parada']].forEach(par => {
      const r = [par[0]];
      for (let d = 1; d <= ndias; d++) { const k = por[d]; r.push(k && k[par[1]] != null ? k[par[1]] / 1440 : ''); }
      r.push('=SUM(B' + (L.length + 1) + ':' + col_(ndias + 1) + (L.length + 1) + ')');
      L.push(r); fmt.push({ r: L.length, t: 'time' });
    });
    const oc = []; ps.forEach(p => (p.paradas || []).forEach(x => oc.push('Dia ' + br_(p.data).slice(0, 5) + ': parou ' + x.de + ' às ' + x.ate + ' – ' + (x.motivo || ''))));
    L.push(['Ocorrências', oc.join('; ') || 'Sem paradas registradas']); L.push(['']);
  });
  const W = ndias + 2;
  const grid = L.map(r => { const x = r.slice(); while (x.length < W) x.push(''); return x; });
  if (ws.getMaxColumns() < W) ws.insertColumnsAfter(ws.getMaxColumns(), W - ws.getMaxColumns());
  ws.getRange(1, 1, grid.length, W).setValues(grid);
  ws.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  fmt.forEach(f => {
    if (f.t === 'grp') ws.getRange(f.r, 1).setFontWeight('bold');
    if (f.t === 'hdr') ws.getRange(f.r, 1, 1, W).setFontWeight('bold').setBackground('#E3EFEE').setHorizontalAlignment('center');
    if (f.t === 'time') { ws.getRange(f.r, 2, 1, W - 1).setNumberFormat('[h]:mm'); ws.getRange(f.r, 1).setFontWeight('bold'); ws.getRange(f.r, W).setFontWeight('bold'); }
  });
  ws.setColumnWidth(1, 95); for (let i = 2; i <= W; i++) ws.setColumnWidth(i, 46); ws.setColumnWidth(W, 62);

  const w2 = ss.insertSheet('Partes diárias');
  const cab = ['Data', 'Prefixo', 'Equipamento', 'Operador', 'Empresa', 'Medidor', 'Inicial', 'Final', 'Diferença', 'Início', 'Término',
               'Jornada (h)', 'Intervalo (h)', 'HD (h)', 'Horas paradas (h)', 'HT – efetivas (h)', 'Abastecimento (L)', 'Paradas (detalhe)', 'Rua / observação'];
  const h = m => m == null ? '' : Math.round(m / 60 * 100) / 100;
  const rows = partes.slice().sort((a, b) => (a.prefixo || '').localeCompare(b.prefixo || '') || a.data.localeCompare(b.data)).map(p => {
    const k = calc_(p, c.intervaloMin);
    return [br_(p.data), p.prefixo, p.descricao, p.operador, p.empresa, p.medidor === 'km' ? 'km' : 'h', p.horIni == null ? '' : p.horIni, p.horFim == null ? '' : p.horFim,
            k.dif == null ? '' : k.dif, p.hIni || '', p.hFim || '', h(k.jornada), h(k.intervalo), h(k.hd), h(k.parada), h(k.efetivo),
            p.abastLitros == null ? '' : p.abastLitros, (p.paradas || []).map(x => x.de + '-' + x.ate + ' ' + (x.motivo || '')).join('; '), p.obs || ''];
  });
  w2.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#E3EFEE');
  if (rows.length) { w2.getRange(2, 1, rows.length, 1).setNumberFormat('@'); w2.getRange(2, 1, rows.length, cab.length).setValues(rows); }
  w2.setFrozenRows(1);
  return exportar_(ss, 'xlsx');
}
function col_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/* ======================= DADOS INICIAIS ======================= */

// Exemplos — apague ou ajuste depois na aba Equipamentos do app.
const EQUIP_INICIAL = [
  { id: 're-01-20260101', prefixo: 'RE 01', descricao: 'RETROESCAVADEIRA', medidor: 'h', operador: 'OPERADOR 1', empresa: 'EMPRESA A', valorMensal: 16000, inicio: '2026-01-01', saida: null },
  { id: 'cb-01-20260101', prefixo: 'CB 01', descricao: 'CAMINHAO BASCULANTE', medidor: 'km', operador: 'MOTORISTA 1', empresa: 'EMPRESA B', valorMensal: 13000, inicio: '2026-01-01', saida: null }
];

const PARTE_INICIAL = null; // opcional: uma parte de exemplo para a base começar preenchida

// Logo da empresa em PNG codificado em base64 (deixe vazio para não usar logo).
const LOGO_B64 = '';