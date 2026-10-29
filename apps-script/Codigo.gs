/**
 * Confirmação de presença em festas — Google Apps Script.
 *
 * Link do convidado:  <URL do app>?e=<código do evento>
 * Painel do admin:    <URL do app>?admin
 *
 * Os dados ficam na própria planilha onde este script está:
 *   aba "Eventos"       → uma linha por festa
 *   aba "Confirmações"  → uma linha por convidado que confirmou
 * As imagens dos convites ficam numa pasta do seu Google Drive.
 */

// ⚠️ Troque antes de publicar. É a senha do painel de admin.
const SENHA_ADMIN = 'troque-esta-senha';

const ABA_EVENTOS = 'Eventos';
const ABA_CONFIRMACOES = 'Confirmações';
const PASTA_CONVITES = 'Convites de festa (confirmação)';
const MAX_ACOMPANHANTES = 30;

// ---------------------------------------------------------------- páginas

function doGet(e) {
  const p = (e && e.parameter) || {};
  const ehAdmin = 'admin' in p;
  const tpl = HtmlService.createTemplateFromFile(ehAdmin ? 'Admin' : 'Convidado');
  tpl.eventoId = String(p.e || '');
  let titulo = 'Confirmação de presença';
  if (ehAdmin) {
    titulo = 'Admin — Confirmações';
  } else if (p.e) {
    const ev = acharEvento_(String(p.e));
    if (ev && ev.titulo) titulo = ev.titulo;
  }
  return tpl.evaluate()
    .setTitle(titulo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Usado nos HTML para incluir o CSS comum: <?!= include_('Estilo'); ?> */
function include_(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

// ---------------------------------------------------------------- convidado

/** Dados públicos de um evento: título, mensagem, imagem do convite. */
function getEvento(id) {
  const ev = acharEvento_(id);
  if (!ev) return null;
  return {
    id: ev.id,
    titulo: ev.titulo,
    mensagem: ev.mensagem,
    aberto: ev.aberto,
    imagem: imagemDataUrl_(ev.imagemId),
  };
}

/** Grava (ou atualiza, se o mesmo nome já confirmou) a presença. */
function confirmar(dados) {
  const ev = acharEvento_(String(dados.evento || ''));
  if (!ev) throw new Error('Evento não encontrado.');
  if (!ev.aberto) throw new Error('As confirmações para este evento estão encerradas.');

  const nome = limparNome_(dados.nome);
  if (nome.length < 2) throw new Error('Informe seu nome.');
  const adultos = inteiro_(dados.adultos);
  const criancas = inteiro_(dados.criancas);

  return comTrava_(() => {
    const sh = aba_(ABA_CONFIRMACOES);
    const linha = acharConfirmacao_(sh, ev.id, nome);
    const valores = [[ev.id, new Date(), seguro_(nome), adultos, criancas]];
    if (linha) {
      sh.getRange(linha, 1, 1, 5).setValues(valores);
      return { atualizado: true };
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues(valores);
    return { atualizado: false };
  });
}

// ---------------------------------------------------------------- admin

function adminEntrar(senha) {
  checar_(senha);
  return { url: ScriptApp.getService().getUrl(), eventos: listarEventos_() };
}

/** Cria (sem id) ou edita (com id) um evento. imagem = dataURL opcional. */
function adminSalvarEvento(senha, dados) {
  checar_(senha);
  const titulo = String(dados.titulo || '').trim();
  if (!titulo) throw new Error('Dê um nome ao evento.');
  const mensagem = String(dados.mensagem || '').trim();
  const aberto = dados.aberto !== false;

  return comTrava_(() => {
    const sh = aba_(ABA_EVENTOS);
    let ev = dados.id ? acharEvento_(dados.id) : null;
    if (dados.id && !ev) throw new Error('Evento não encontrado.');

    let imagemId = ev ? ev.imagemId : '';
    if (dados.imagem) {
      const novo = salvarImagem_(dados.imagem, titulo);
      if (imagemId) {
        try { DriveApp.getFileById(imagemId).setTrashed(true); } catch (e) { /* já apagada */ }
      }
      imagemId = novo;
    }

    if (ev) {
      sh.getRange(ev.linha, 2, 1, 3).setValues([[seguro_(titulo), seguro_(mensagem), imagemId]]);
      sh.getRange(ev.linha, 6).setValue(aberto);
      return ev.id;
    }
    const id = novoId_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, 6)
      .setValues([[id, seguro_(titulo), seguro_(mensagem), imagemId, new Date(), aberto]]);
    return id;
  });
}

/** Dados completos do evento para edição + confirmações. */
function adminEvento(senha, id) {
  checar_(senha);
  const ev = getEvento(id);
  if (!ev) throw new Error('Evento não encontrado.');
  ev.confirmacoes = listarConfirmacoes_(id);
  return ev;
}

function adminExcluirConfirmacao(senha, eventoId, nome) {
  checar_(senha);
  return comTrava_(() => {
    const sh = aba_(ABA_CONFIRMACOES);
    const linha = acharConfirmacao_(sh, eventoId, nome);
    if (linha) sh.deleteRow(linha);
    return listarConfirmacoes_(eventoId);
  });
}

function adminExcluirEvento(senha, id) {
  checar_(senha);
  return comTrava_(() => {
    const ev = acharEvento_(id);
    if (!ev) return listarEventos_();
    if (ev.imagemId) {
      try { DriveApp.getFileById(ev.imagemId).setTrashed(true); } catch (e) { /* já apagada */ }
    }
    const sc = aba_(ABA_CONFIRMACOES);
    const vals = sc.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === id) sc.deleteRow(i + 1);
    }
    aba_(ABA_EVENTOS).deleteRow(ev.linha);
    return listarEventos_();
  });
}

// ---------------------------------------------------------------- planilha

/**
 * A planilha onde ficam os dados. Se o script foi criado dentro de uma planilha
 * (Extensões → Apps Script), usa ela. Se foi criado avulso em script.google.com,
 * cria uma planilha "Confirmações de festas" no seu Drive e passa a usar essa.
 */
function planilha_() {
  const ativa = SpreadsheetApp.getActive();
  if (ativa) return ativa;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PLANILHA_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* apagada: cria outra */ }
  }
  const nova = SpreadsheetApp.create('Confirmações de festas');
  props.setProperty('PLANILHA_ID', nova.getId());
  return nova;
}

function aba_(nome) {
  const ss = planilha_();
  let sh = ss.getSheetByName(nome);
  if (sh) return sh;
  sh = ss.insertSheet(nome);
  const cab = nome === ABA_EVENTOS
    ? ['ID', 'Título', 'Mensagem', 'Imagem (Drive)', 'Criado em', 'Aberto']
    : ['Evento', 'Data/hora', 'Nome', 'Adultos acompanhantes', 'Crianças acompanhantes (até 12)'];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function listarEventos_() {
  const vals = aba_(ABA_EVENTOS).getDataRange().getValues();
  const contagem = {};
  aba_(ABA_CONFIRMACOES).getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[0]);
    const c = contagem[id] || (contagem[id] = { confirmacoes: 0, pessoas: 0 });
    c.confirmacoes++;
    c.pessoas += 1 + Number(r[3] || 0) + Number(r[4] || 0);
  });
  return vals.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    titulo: String(r[1]),
    aberto: r[5] !== false,
    criado: r[4] instanceof Date ? r[4].toISOString() : '',
    confirmacoes: (contagem[r[0]] || {}).confirmacoes || 0,
    pessoas: (contagem[r[0]] || {}).pessoas || 0,
  })).reverse();
}

function acharEvento_(id) {
  if (!id) return null;
  const vals = aba_(ABA_EVENTOS).getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[0]) === id) {
      return {
        linha: i + 1, id: String(r[0]), titulo: String(r[1]), mensagem: String(r[2]),
        imagemId: String(r[3] || ''), aberto: r[5] !== false,
      };
    }
  }
  return null;
}

function listarConfirmacoes_(eventoId) {
  return aba_(ABA_CONFIRMACOES).getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === eventoId)
    .map(r => ({
      nome: String(r[2]),
      adultos: Number(r[3] || 0),
      criancas: Number(r[4] || 0),
      quando: r[1] instanceof Date ? r[1].toISOString() : '',
    }));
}

function acharConfirmacao_(sh, eventoId, nome) {
  const chave = normalizar_(nome);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === eventoId && normalizar_(vals[i][2]) === chave) return i + 1;
  }
  return 0;
}

// ---------------------------------------------------------------- imagens

function pastaConvites_() {
  const it = DriveApp.getFoldersByName(PASTA_CONVITES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_CONVITES);
}

function salvarImagem_(dataUrl, titulo) {
  const m = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i.exec(String(dataUrl));
  if (!m) throw new Error('Imagem inválida.');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'convite - ' + titulo + '.' + ext);
  return pastaConvites_().createFile(blob).getId();
}

function imagemDataUrl_(id) {
  if (!id) return null;
  try {
    const b = DriveApp.getFileById(id).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) {
    return null;
  }
}

// ---------------------------------------------------------------- diagnóstico

/**
 * Rode esta função no editor (escolha "diagnostico" no menu ao lado de ▷ Executar)
 * e veja o resultado em "Registro de execução". Ela também pede as autorizações
 * do Google, se faltar alguma.
 */
function diagnostico() {
  const senha = String(SENHA_ADMIN).trim();
  Logger.log('Senha padrão ainda? ' + (senha === 'troque-esta-senha' ? 'SIM, troque!' : 'não'));
  Logger.log('Tamanho da senha: ' + senha.length + ' caracteres (sem contar espaços nas pontas)');
  const ss = planilha_();
  Logger.log('Planilha usada: ' + ss.getUrl());
  Logger.log('Eventos cadastrados: ' + listarEventos_().length);
  Logger.log('Pasta dos convites: ' + pastaConvites_().getUrl());
  Logger.log('URL publicada: ' + (ScriptApp.getService().getUrl() || '(nenhuma implantação ainda)'));
  Logger.log('Tudo certo do lado do servidor. Lembre: depois de mudar o código, publique uma nova versão.');
}

// ---------------------------------------------------------------- utilidades

function checar_(senha) {
  const certa = String(SENHA_ADMIN).trim();
  if (certa === 'troque-esta-senha') {
    throw new Error('A senha ainda é a padrão. Troque SENHA_ADMIN no Codigo.gs e publique uma nova versão ' +
      '(Implantar → Gerenciar implantações → ✏️ → Nova versão).');
  }
  if (String(senha || '').trim() !== certa) {
    throw new Error('Senha incorreta. Se você trocou a senha há pouco, confira se publicou uma nova versão ' +
      '(Implantar → Gerenciar implantações → ✏️ → Nova versão).');
  }
}

function comTrava_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function novoId_() {
  const letras = 'abcdefghjkmnpqrstuvwxyz23456789';
  let id = '';
  for (let i = 0; i < 6; i++) id += letras[Math.floor(Math.random() * letras.length)];
  return acharEvento_(id) ? novoId_() : id;
}

function inteiro_(v) {
  const n = Math.floor(Number(v));
  return isFinite(n) && n > 0 ? Math.min(n, MAX_ACOMPANHANTES) : 0;
}

function limparNome_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().slice(0, 120);
}

function normalizar_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Evita que um texto começando com = + - @ vire fórmula na planilha. */
function seguro_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
