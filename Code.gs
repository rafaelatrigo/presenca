/**
 * Lista de Presença – Carozzo
 * Cole este código em: Planilha > Extensões > Apps Script
 * Depois: Implantar > Nova implantação > App da Web
 *   - Executar como: Eu
 *   - Quem pode acessar: Qualquer pessoa
 */

// Código que a recepção digita na primeira vez que abre a página.
// Troque antes do evento e passe só para a equipe da recepção.
const CODIGO_ACESSO = 'carozzo2026';

// Nomes das colunas (linha 1 da planilha). O script localiza cada coluna pelo título,
// então você pode reordenar ou incluir colunas novas na planilha sem mexer aqui.
const COL = {
  id: 'ID',
  nome: 'Nome',
  empresa: 'Empresa',
  categoria: 'Categoria',
  status: 'Status',
  chegada: 'Chegada',
  por: 'Confirmado por'
};
const PRESENTE = 'Presente';

function doGet(e) {
  const p = e.parameter || {};
  try {
    if (p.codigo !== CODIGO_ACESSO) return saida_({ ok: false, erro: 'codigo' });
    switch (p.acao) {
      case 'listar':     return saida_({ ok: true, convidados: listar_() });
      case 'confirmar':  return saida_(alterar_(p.id, true, p.por));
      case 'desfazer':   return saida_(alterar_(p.id, false, p.por));
      case 'adicionar':  return saida_(adicionar_(p.nome, p.empresa, p.por));
      default:           return saida_({ ok: false, erro: 'acao' });
    }
  } catch (err) {
    return saida_({ ok: false, erro: String(err) });
  }
}

function aba_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function mapa_(sh) {
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const m = {};
  Object.keys(COL).forEach(k => {
    const i = cab.indexOf(COL[k]);
    if (i < 0) throw new Error('Coluna não encontrada: ' + COL[k]);
    m[k] = i;
  });
  return m;
}

function listar_() {
  const sh = aba_();
  const m = mapa_(sh);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, n, sh.getLastColumn()).getValues()
    .filter(r => r[m.nome] !== '')
    .map(r => ({
      id: String(r[m.id]),
      nome: r[m.nome],
      empresa: r[m.empresa],
      categoria: r[m.categoria],
      presente: r[m.status] === PRESENTE,
      chegada: r[m.chegada] instanceof Date
        ? Utilities.formatDate(r[m.chegada], tz, 'HH:mm') : String(r[m.chegada] || ''),
      por: r[m.por]
    }));
}

function alterar_(id, presente, por) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita conflito quando 2 aparelhos clicam ao mesmo tempo
  try {
    const sh = aba_();
    const m = mapa_(sh);
    const ids = sh.getRange(2, m.id + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(id));
    if (i < 0) return { ok: false, erro: 'nao_encontrado' };
    const linha = i + 2;
    const atual = sh.getRange(linha, m.status + 1).getValue() === PRESENTE;
    if (presente && atual) {
      // Outro aparelho já confirmou: não sobrescreve horário nem responsável
      return { ok: true, jaConfirmado: true, convidados: listar_() };
    }
    sh.getRange(linha, m.status + 1).setValue(presente ? PRESENTE : '');
    sh.getRange(linha, m.chegada + 1).setValue(presente ? new Date() : '');
    sh.getRange(linha, m.por + 1).setValue(presente ? (por || '') : '');
    SpreadsheetApp.flush();
    return { ok: true, convidados: listar_() };
  } finally {
    lock.releaseLock();
  }
}

function adicionar_(nome, empresa, por) {
  if (!nome) return { ok: false, erro: 'nome' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba_();
    const m = mapa_(sh);
    const ids = sh.getLastRow() > 1
      ? sh.getRange(2, m.id + 1, sh.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0) : [0];
    const novoId = Math.max.apply(null, ids) + 1;
    const linha = new Array(sh.getLastColumn()).fill('');
    linha[m.id] = novoId;
    linha[m.nome] = nome;
    linha[m.empresa] = empresa || '';
    linha[m.categoria] = 'Não listado';
    linha[m.status] = PRESENTE;
    linha[m.chegada] = new Date();
    linha[m.por] = por || '';
    sh.appendRow(linha);
    SpreadsheetApp.flush();
    return { ok: true, convidados: listar_() };
  } finally {
    lock.releaseLock();
  }
}

function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
