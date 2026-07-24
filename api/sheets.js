const { google } = require('googleapis');

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function lerAba(sheets, aba) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${aba}!A:Z`,
    });
    const rows = res.data.values || [];
    if (rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] || ''; });
      return obj;
    });
  } catch (e) {
    return [];
  }
}

async function garantirCabecalhos(sheets) {
  const abas = {
    usuarios:      ['id','nome','area','email','senha','tipo','foto','ativo','criado_em'],
    okr_companhia: ['id','nome','descricao','responsavel','data_inicio','data_fim','progresso','saude','status','criado_em'],
    okr_area:      ['id','okr_companhia_id','nome','descricao','area','responsavel','data_inicio','data_fim','progresso','saude','status','criado_em'],
    key_results:   ['id','okr_id','tipo','nome','descricao','responsavel','data_inicio','data_fim','progresso','saude','indicador','meta','atual','unidade','criado_em'],
    iniciativas:   ['id','kr_id','nome','descricao','responsavel','data_inicio','data_fim','status','comentario','pct','criado_em'],
    aprovacoes:    ['id','tipo','nome','area','solicitante','data','status','aprovado_por','data_aprovacao'],
    snapshots:     ['id','mes','ano','area','okr_id','kr_id','progresso','saude','criado_em'],
  };

  for (const [aba, headers] of Object.entries(abas)) {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${aba}!A1:Z1`,
      });
      const existing = res.data.values?.[0] || [];
      if (existing.length === 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `${aba}!A1`,
          valueInputOption: 'RAW',
          requestBody: { values: [headers] },
        });
      }
    } catch (e) {
      // aba pode não existir ainda
    }
  }
}

async function proximoId(sheets, aba) {
  const dados = await lerAba(sheets, aba);
  if (dados.length === 0) return 1;
  const ids = dados.map(r => parseInt(r.id) || 0);
  return Math.max(...ids) + 1;
}

async function adicionarLinha(sheets, aba, valores) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${aba}!A:A`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [valores] },
  });
}

async function atualizarLinha(sheets, aba, id, novosValores, headers) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${aba}!A:Z`,
  });
  const rows = res.data.values || [];
  const rowIndex = rows.findIndex((r, i) => i > 0 && r[0] == id);
  if (rowIndex === -1) return false;
  const sheetRow = rowIndex + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${aba}!A${sheetRow}`,
    valueInputOption: 'RAW',
    requestBody: { values: [novosValores] },
  });
  return true;
}

function agora() {
  return new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const sheets = await getSheets();
    await garantirCabecalhos(sheets);

    const { acao, aba } = req.query;
    const body = req.body || {};

    // ── LISTAR ──────────────────────────────────────────
    if (req.method === 'GET' && acao === 'listar') {
      const dados = await lerAba(sheets, aba);
      return res.status(200).json({ ok: true, dados });
    }

    // ── CRIAR ────────────────────────────────────────────
    if (req.method === 'POST' && acao === 'criar') {
      const id = await proximoId(sheets, aba);
      const ts = agora();

      let linha;
      if (aba === 'usuarios') {
        linha = [id, body.nome, body.area, body.email, body.senha, body.tipo, body.foto || '', 'sim', ts];
      } else if (aba === 'okr_companhia') {
        linha = [id, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, 0, 'green', 'Pendente', ts];
      } else if (aba === 'okr_area') {
        linha = [id, body.okr_companhia_id, body.nome, body.descricao, body.area, body.responsavel, body.data_inicio, body.data_fim, 0, 'green', 'Pendente', ts];
      } else if (aba === 'key_results') {
        linha = [id, body.okr_id, body.tipo, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, 0, 'green', body.indicador || '', body.meta || '', body.atual || 0, body.unidade || '', ts];
      } else if (aba === 'iniciativas') {
        linha = [id, body.kr_id, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, body.status || 'Não iniciada', body.comentario || '', 0, ts];
      } else if (aba === 'aprovacoes') {
        linha = [id, body.tipo, body.nome, body.area, body.solicitante, ts, 'Pendente', '', ''];
      } else if (aba === 'snapshots') {
        const d = new Date();
        linha = [id, d.getMonth() + 1, d.getFullYear(), body.area, body.okr_id || '', body.kr_id || '', body.progresso, body.saude, ts];
      }

      if (!linha) return res.status(400).json({ ok: false, erro: 'Aba não reconhecida' });
      await adicionarLinha(sheets, aba, linha);
      return res.status(200).json({ ok: true, id });
    }

    // ── ATUALIZAR ────────────────────────────────────────
    if (req.method === 'PUT' && acao === 'atualizar') {
      const { id } = body;
      let linha;

      if (aba === 'key_results') {
        const atual = parseFloat(body.atual) || 0;
        const meta = parseFloat(body.meta) || 1;
        let prog = body.indicador === 'crescente'
          ? Math.min(Math.round((atual / meta) * 100), 100)
          : Math.max(Math.round((1 - atual / meta) * 100), 0);
        linha = [id, body.okr_id, body.tipo, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, prog, body.saude, body.indicador, meta, atual, body.unidade, body.criado_em];
      } else if (aba === 'iniciativas') {
        const pctMap = { 'Concluída': 100, 'Concluída com atraso': 100, 'Em andamento': 50, 'Não iniciada': 0, 'Cancelada': 0 };
        const pct = pctMap[body.status] ?? 0;
        linha = [id, body.kr_id, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, body.status, body.comentario, pct, body.criado_em];
      } else if (aba === 'aprovacoes') {
        linha = [id, body.tipo, body.nome, body.area, body.solicitante, body.data, body.status, body.aprovado_por || '', body.data_aprovacao || agora()];
      } else if (aba === 'usuarios') {
        linha = [id, body.nome, body.area, body.email, body.senha, body.tipo, body.foto || '', body.ativo || 'sim', body.criado_em];
      } else if (aba === 'okr_companhia') {
        linha = [id, body.nome, body.descricao, body.responsavel, body.data_inicio, body.data_fim, body.progresso, body.saude, body.status, body.criado_em];
      } else if (aba === 'okr_area') {
        linha = [id, body.okr_companhia_id, body.nome, body.descricao, body.area, body.responsavel, body.data_inicio, body.data_fim, body.progresso, body.saude, body.status, body.criado_em];
      }

      if (!linha) return res.status(400).json({ ok: false, erro: 'Aba não reconhecida' });
      const ok = await atualizarLinha(sheets, aba, id, linha);
      return res.status(200).json({ ok });
    }

    return res.status(400).json({ ok: false, erro: 'Ação não reconhecida' });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, erro: err.message });
  }
};
