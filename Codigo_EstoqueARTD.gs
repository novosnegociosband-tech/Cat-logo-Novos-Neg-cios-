/**
 * Estoque ARTD — fotos e planilha no Google Drive
 *
 * Guarda tudo numa pasta do seu Drive:
 *   - as fotos dos materiais (uma imagem por código)
 *   - fotos.json   → qual foto é de qual código
 *   - estoque.json → a última planilha enviada pelo site
 *
 * CONFIGURE AS DUAS LINHAS ABAIXO e publique como App da Web.
 */
const PASTA_ID = 'COLE_AQUI_O_ID_DA_PASTA';   // parte final do link da pasta: drive.google.com/drive/folders/ESTE_ID
const SENHA    = 'troque-esta-senha';          // a mesma que você digita no site para editar

/* ---------- leitura (qualquer pessoa com o link do site) ---------- */
function doGet(e) {
  try {
    const pasta = DriveApp.getFolderById(PASTA_ID);
    const idx = lerJson_(pasta, 'fotos.json') || {};
    const fotos = {};
    Object.keys(idx).forEach(function (cod) {
      const v = idx[cod];
      if (v.r) fotos[cod] = { removed: true };
      else if (v.f) fotos[cod] = { url: urlFoto_(v.f, v.t) };
    });
    const estoque = lerTexto_(pasta, 'estoque.json');
    return saida_({ ok: true, fotos: fotos, estoque: estoque || null });
  } catch (err) {
    return saida_({ ok: false, erro: 'Erro no Drive: ' + err.message });
  }
}

/* ---------- escrita (exige a senha) ---------- */
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); }
  catch (err) { return saida_({ ok: false, erro: 'Pedido inválido.' }); }
  if (String(b.pin || '') !== SENHA) return saida_({ ok: false, erro: 'Senha incorreta.' });

  const trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    const pasta = DriveApp.getFolderById(PASTA_ID);
    if (b.action === 'ping') return saida_({ ok: true });
    if (b.action === 'foto') return saida_(foto_(pasta, b));
    if (b.action === 'estoque') return saida_(estoque_(pasta, b));
    return saida_({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return saida_({ ok: false, erro: 'Erro no Drive: ' + err.message });
  } finally {
    trava.releaseLock();
  }
}

function foto_(pasta, b) {
  const cod = String(b.id || '').trim();
  if (!cod) return { ok: false, erro: 'Código do material vazio.' };
  const idx = lerJson_(pasta, 'fotos.json') || {};
  const antigo = idx[cod] && idx[cod].f;

  if (b.revert) {                      // volta para a foto que já vem no site
    delete idx[cod];
  } else if (b.removed) {              // deixa o item sem foto
    idx[cod] = { r: true, t: Date.now() };
  } else {                             // foto nova
    if (!b.data) return { ok: false, erro: 'Foto vazia.' };
    const bytes = Utilities.base64Decode(b.data);
    if (bytes.length > 4 * 1024 * 1024) return { ok: false, erro: 'Foto grande demais (máx. 4 MB).' };
    const arq = pasta.createFile(Utilities.newBlob(bytes, b.mime || 'image/jpeg', 'material_' + cod + '_' + Date.now() + '.jpg'));
    arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    idx[cod] = { f: arq.getId(), t: Date.now() };
  }
  gravarTexto_(pasta, 'fotos.json', JSON.stringify(idx));
  if (antigo && (!idx[cod] || idx[cod].f !== antigo)) {
    try { DriveApp.getFileById(antigo).setTrashed(true); } catch (err) {}
  }
  return { ok: true, url: idx[cod] && idx[cod].f ? urlFoto_(idx[cod].f, idx[cod].t) : null };
}

function estoque_(pasta, b) {
  const json = String(b.json || '');
  if (!json) {                         // "voltar para a versão original" no site
    const it = pasta.getFilesByName('estoque.json');
    while (it.hasNext()) it.next().setTrashed(true);
    return { ok: true };
  }
  const d = JSON.parse(json);
  if (!d.cad || !d.mv) return { ok: false, erro: 'Planilha sem as abas Cadastro e Movimentações.' };
  gravarTexto_(pasta, 'estoque.json', json);
  return { ok: true, materiais: d.cad.length, movimentacoes: d.mv.length };
}

/* ---------- utilitários ---------- */
function urlFoto_(fileId, t) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000&v=' + (t || 0);
}
function lerTexto_(pasta, nome) {
  const it = pasta.getFilesByName(nome);
  return it.hasNext() ? it.next().getBlob().getDataAsString('UTF-8') : '';
}
function lerJson_(pasta, nome) {
  const s = lerTexto_(pasta, nome);
  try { return s ? JSON.parse(s) : null; } catch (err) { return null; }
}
function gravarTexto_(pasta, nome, texto) {
  const it = pasta.getFilesByName(nome);
  if (it.hasNext()) it.next().setContent(texto);
  else pasta.createFile(nome, texto, MimeType.PLAIN_TEXT);
}
function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Rode esta função uma vez pelo editor para autorizar o acesso ao Drive. */
function autorizar() {
  const pasta = DriveApp.getFolderById(PASTA_ID);
  Logger.log('Pasta encontrada: ' + pasta.getName());
}
