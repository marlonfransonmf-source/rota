/** Rota Visitas · roteiro de visita a centros de serviço de paletes
 *  Dados numa planilha criada automaticamente, fotos e relatórios no Drive, e-mail pela conta de quem usa.
 */
const APP = 'Rota Visitas';

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.email = Session.getActiveUser().getEmail() || '';
  return t.evaluate().setTitle(APP)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function props_() { return PropertiesService.getScriptProperties(); }

function planilha_() {
  const p = props_(); let id = p.getProperty('SHEET');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(APP + ' · Base');
  const v = ss.getSheets()[0]; v.setName('Visitas');
  v.appendRow(['ID', 'Data', 'Centro', 'Visitante', 'Nota %', 'Itens NC', 'Relatório PDF', 'JSON']); v.setFrozenRows(1);
  const c = ss.insertSheet('Centros'); c.appendRow(['Centro', 'Cidade', 'Responsável', 'E-mail']); c.setFrozenRows(1);
  [['CS Cajamar', 'Cajamar-SP', '', ''], ['CS Contagem', 'Contagem-MG', '', ''], ['CS Recife', 'Recife-PE', '', '']].forEach(r => c.appendRow(r));
  const a = ss.insertSheet('Ações'); a.appendRow(['ID', 'Visita', 'Centro', 'Item', 'O que corrigir', 'Responsável', 'Prazo', 'Status']); a.setFrozenRows(1);
  p.setProperty('SHEET', ss.getId());
  pasta_().addFile(DriveApp.getFileById(ss.getId()));
  return ss;
}

function pasta_() {
  const p = props_(); const id = p.getProperty('PASTA');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(APP);
  p.setProperty('PASTA', f.getId()); return f;
}

function carregar() {
  const ss = planilha_();
  const cen = ss.getSheetByName('Centros').getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ nome: r[0], cidade: r[1], resp: r[2], email: r[3] }));
  const vis = ss.getSheetByName('Visitas').getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ id: r[0], data: Utilities.formatDate(new Date(r[1]), 'America/Sao_Paulo', 'dd/MM/yyyy'), centro: r[2], visitante: r[3], nota: r[4], nc: r[5], pdf: r[6] })).reverse().slice(0, 50);
  const ac = ss.getSheetByName('Ações').getDataRange().getValues().slice(1).filter(r => r[0] && r[7] !== 'Concluída')
    .map(r => ({ id: r[0], centro: r[2], item: r[3], txt: r[4], resp: r[5], prazo: r[6] ? Utilities.formatDate(new Date(r[6]), 'America/Sao_Paulo', 'dd/MM') : '' }));
  return JSON.stringify({ centros: cen, visitas: vis, acoes: ac, planilha: ss.getUrl() });
}

function concluirAcao(id) {
  const sh = planilha_().getSheetByName('Ações');
  const f = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  if (f) sh.getRange(f.getRow(), 8).setValue('Concluída');
  return true;
}

function salvarVisita(json) {
  const v = JSON.parse(json);
  const ss = planilha_(), pasta = pasta_();
  const id = 'V' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMdd-HHmmss');
  const fotos = {};
  Object.keys(v.fotos || {}).forEach(k => {
    const b = Utilities.newBlob(Utilities.base64Decode(v.fotos[k]), 'image/jpeg', id + '_' + k + '.jpg');
    fotos[k] = pasta.createFile(b);
  });
  // Relatório (produto final)
  const doc = DocumentApp.create('Relatório de visita · ' + v.centro + ' · ' + v.data);
  const b = doc.getBody(); b.setMarginTop(36).setMarginBottom(36);
  b.appendParagraph('Relatório de visita').setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph(v.centro + ' · ' + v.data + ' · visitante: ' + v.visitante);
  b.appendParagraph('Nota geral: ' + v.nota + '%  ·  ' + v.nc.length + ' não conformidades').setBold(true);
  b.appendParagraph('Resumo por área').setHeading(DocumentApp.ParagraphHeading.HEADING1).setBold(false);
  const tab = [['Área', 'Nota', 'NC']].concat(v.secoes.map(s => [s.nome, s.nota == null ? '-' : s.nota + '%', String(s.nc)]));
  b.appendTable(tab);
  if (v.nc.length) {
    b.appendParagraph('Plano de ação').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    b.appendTable([['Item', 'O que corrigir', 'Responsável', 'Prazo']].concat(v.nc.map(n => [n.item, n.obs || '-', n.resp || '-', n.prazo || '-'])));
  }
  b.appendParagraph('Detalhe da visita').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  v.secoes.forEach(s => {
    b.appendParagraph(s.nome).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s.itens.forEach(it => {
      const st = { c: 'Conforme', nc: 'NÃO CONFORME', na: 'N/A' }[it.v] || 'Não avaliado';
      b.appendListItem(it.txt + ': ' + st + (it.obs ? ' · ' + it.obs : ''));
      if (fotos[it.id]) { try { b.appendImage(fotos[it.id].getBlob()).setWidth(240).setHeight(180); } catch (e) {} }
    });
  });
  if (v.obsGeral) { b.appendParagraph('Observações gerais').setHeading(DocumentApp.ParagraphHeading.HEADING1); b.appendParagraph(v.obsGeral); }
  doc.saveAndClose();
  const pdf = pasta.createFile(doc.getAs('application/pdf'));
  pasta.addFile(DriveApp.getFileById(doc.getId())); DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));
  ss.getSheetByName('Visitas').appendRow([id, new Date(), v.centro, v.visitante, v.nota, v.nc.length, pdf.getUrl(), json.length < 45000 ? JSON.stringify(Object.assign({}, v, { fotos: {} })) : '']);
  const ac = ss.getSheetByName('Ações');
  v.nc.forEach((n, i) => ac.appendRow([id + '-' + (i + 1), id, v.centro, n.item, n.obs || '', n.resp || '', n.prazo ? new Date(n.prazo + 'T12:00:00') : '', 'Aberta']));
  const para = [Session.getActiveUser().getEmail()].concat(v.copia ? [v.copia] : []).filter(Boolean).join(',');
  if (para) MailApp.sendEmail({ to: para, subject: 'Relatório de visita · ' + v.centro + ' · ' + v.nota + '%',
    htmlBody: '<p>Visita ao <b>' + v.centro + '</b> em ' + v.data + '. Nota <b>' + v.nota + '%</b>, ' + v.nc.length + ' não conformidades.</p><p>Relatório em anexo.</p>',
    attachments: [pdf.getAs('application/pdf')], name: APP });
  return JSON.stringify({ id: id, pdf: pdf.getUrl() });
}

/** Rode uma vez pelo editor para criar a planilha e a pasta e autorizar tudo. */
function instalar() {
  const ss = planilha_();
  Logger.log('Pronto. Planilha: ' + ss.getUrl() + ' · Pasta: ' + pasta_().getUrl());
}
