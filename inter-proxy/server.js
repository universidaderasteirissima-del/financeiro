/**
 * Rasteirissima — Proxy Banco Inter Cobranças (API V3)
 *
 * Suporta 3 contas Inter distintas (brilhante, marketing, franchising).
 * O frontend envia `conta: 'brilhante'|'marketing'|'franchising'` no body
 * e o proxy usa as credenciais correspondentes.
 *
 * COMO USAR (desenvolvimento local):
 *   1. Copie .env.example para .env e preencha as 3 contas
 *   2. npm install
 *   3. node server.js
 *
 * DEPLOY (Railway ou Render — plano gratuito):
 *   - Suba esta pasta como repositório GitHub
 *   - Configure todas as variáveis de ambiente (ver .env.example)
 *   - Cole a URL gerada em Configurações → "URL servidor Inter" no sistema
 */

const express = require('express');
const https   = require('https');

const {
  ALLOWED_ORIGIN,
  PORT = 3000,
} = process.env;

// ── Configuração das 3 contas Inter ─────────────────────────────────────────
const CONTAS = {
  brilhante:   buildContaCfg('BRI'),
  marketing:   buildContaCfg('MKT'),
  franchising: buildContaCfg('FRAN'),
};


function buildContaCfg(prefix) {
  return {
    clientId:     process.env[`INTER_${prefix}_CLIENT_ID`]     || '',
    clientSecret: process.env[`INTER_${prefix}_CLIENT_SECRET`] || '',
    certPem:      (process.env[`INTER_${prefix}_CERT_PEM`]     || '').replace(/\\n/g, '\n'),
    keyPem:       (process.env[`INTER_${prefix}_KEY_PEM`]      || '').replace(/\\n/g, '\n'),
    conta:        process.env[`INTER_${prefix}_CONTA`]         || '',
    token:        null,
    tokenExp:     0,
  };
}

const contasValidas = Object.entries(CONTAS).filter(([, c]) => c.clientId);
if (!contasValidas.length) {
  console.error('Nenhuma conta Inter configurada. Defina INTER_BRI_CLIENT_ID, INTER_MKT_CLIENT_ID ou INTER_FRAN_CLIENT_ID.');
  process.exit(1);
}
if (!ALLOWED_ORIGIN) {
  console.warn('ALLOWED_ORIGIN não definido — CORS aberto (não recomendado em produção)');
}

const INTER_BASE = 'https://cdpj.partners.bancointer.com.br';

// ── OAuth2: token por conta com cache de 55 min ───────────────────────────────
async function getToken(cfg) {
  if (cfg.token && Date.now() < cfg.tokenExp) return cfg.token;

  const body = `grant_type=client_credentials`
    + `&client_id=${encodeURIComponent(cfg.clientId)}`
    + `&client_secret=${encodeURIComponent(cfg.clientSecret)}`
    + `&scope=boleto-cobranca.read%20boleto-cobranca.write`;

  const result = await interRequest('/oauth/v2/token', {
    method: 'POST',
    cfg,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });

  if (result.status !== 200) {
    throw new Error(`Inter OAuth falhou (${result.status}): ${result.body}`);
  }
  const json = JSON.parse(result.body);
  cfg.token = json.access_token;
  cfg.tokenExp = Date.now() + (json.expires_in - 60) * 1000;
  return cfg.token;
}

// ── Requisição HTTPS com mTLS ─────────────────────────────────────────────────
function interRequest(path, { method = 'GET', cfg, headers = {}, body = null, token }) {
  return new Promise((resolve, reject) => {
    const url = new URL(INTER_BASE + path);
    const buf = body ? Buffer.from(body, 'utf8') : null;

    const reqHeaders = {
      ...headers,
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(buf ? { 'Content-Length': String(buf.length) } : {}),
      ...(cfg.conta ? { 'x-conta-corrente': cfg.conta } : {}),
    };

    const tlsOpts = {};
    if (cfg.certPem) tlsOpts.cert = cfg.certPem;
    if (cfg.keyPem)  tlsOpts.key  = cfg.keyPem;

    const req = https.request(
      {
        hostname: url.hostname,
        path: url.pathname + url.search,
        method,
        headers: reqHeaders,
        ...tlsOpts,
      },
      (res) => {
        let data = '';
        res.on('data', c => { data += c; });
        res.on('end', () => resolve({ status: res.statusCode, body: data }));
      }
    );
    req.on('error', reject);
    if (buf) req.write(buf);
    req.end();
  });
}

// Aguarda N ms
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ── Servidor Express ──────────────────────────────────────────────────────────
const app = express();
app.use(express.json());

app.use((req, res, next) => {
  const origin = req.headers.origin;
  const allowed = ALLOWED_ORIGIN || '*';
  if (!ALLOWED_ORIGIN || origin === ALLOWED_ORIGIN) {
    res.setHeader('Access-Control-Allow-Origin', allowed);
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  if (req.method === 'OPTIONS') { res.sendStatus(204); return; }
  next();
});

// GET /ping
app.get('/ping', (_req, res) => {
  res.json({
    ok: true,
    contas: Object.entries(CONTAS)
      .filter(([, c]) => c.clientId)
      .map(([nome]) => nome),
  });
});

// POST /boleto — cria cobrança V3 (Boleto + Pix)
//
// Body esperado:
// {
//   conta: 'brilhante' | 'marketing' | 'franchising',
//   sacado: { cnpjCpf, nome, email, cep, endereco, numero, complemento, bairro, cidade, uf, telefone },
//   valor: 1234.56,
//   vencimento: 'AAAA-MM-DD',
//   descricao: 'Boleto Marketing — NF 123 — Loja XYZ',
//   seuNumero: 'REF123' // opcional, max 15 chars; gerado automaticamente se omitido
// }
app.post('/boleto', async (req, res) => {
  try {
    const { conta, sacado, valor, vencimento, descricao, seuNumero } = req.body;

    if (!conta || !CONTAS[conta]) {
      return res.status(400).json({ error: `Conta inválida: "${conta}". Use brilhante, marketing ou franchising.` });
    }
    if (!sacado || !sacado.cnpjCpf || !valor || !vencimento) {
      return res.status(400).json({ error: 'Campos obrigatórios: sacado.cnpjCpf, valor, vencimento' });
    }

    const cfg = CONTAS[conta];
    if (!cfg.clientId) {
      return res.status(503).json({ error: `Conta "${conta}" não está configurada neste servidor.` });
    }

    const token = await getToken(cfg);

    const cnpjCpfLimpo = sacado.cnpjCpf.replace(/\D/g, '');

    // seuNumero: referência interna, máx 15 chars
    const refNum = (seuNumero || String(Date.now()).slice(-15)).slice(0, 15);

    const payload = JSON.stringify({
      seuNumero:      refNum,
      valorNominal:   Number(valor),
      dataVencimento: vencimento,
      numDiasAgenda:  60,
      pagador: {
        cpfCnpj:     cnpjCpfLimpo,
        tipoPessoa:  cnpjCpfLimpo.length === 11 ? 'FISICA' : 'JURIDICA',
        nome:        sacado.nome,
        email:       sacado.email        || '',
        telefone:    (sacado.telefone    || '').replace(/\D/g, ''),
        endereco:    sacado.endereco     || '',
        numero:      sacado.numero       || 'S/N',
        complemento: sacado.complemento  || '',
        bairro:      sacado.bairro       || '',
        cidade:      sacado.cidade       || '',
        uf:          (sacado.uf || '').toUpperCase().slice(0, 2),
        cep:         sacado.cep.replace(/\D/g, ''),
      },
      mensagem: {
        linha1: (descricao || 'Rasteirissima').slice(0, 50),
      },
    });

    // Passo 1: POST emite a cobrança (assíncrono — retorna codigoSolicitacao)
    const createRes = await interRequest('/cobranca/v3/cobrancas', {
      method: 'POST',
      cfg,
      token,
      headers: { 'Content-Type': 'application/json' },
      body: payload,
    });

    if (createRes.status !== 200 && createRes.status !== 201 && createRes.status !== 202) {
      return res.status(502).json({
        error: `Inter respondeu ${createRes.status}`,
        detail: createRes.body,
      });
    }

    const createData = JSON.parse(createRes.body);
    const codigoSolicitacao = createData.codigoSolicitacao || createData.nossoNumero || '';

    if (!codigoSolicitacao) {
      return res.status(502).json({ error: 'Inter não retornou codigoSolicitacao', detail: createRes.body });
    }

    // Passo 2: polling — aguarda o boleto ficar disponível (até ~10s)
    let cobranca = null;
    for (let i = 0; i < 6; i++) {
      await sleep(i === 0 ? 1500 : 1500);
      const getRes = await interRequest(`/cobranca/v3/cobrancas/${codigoSolicitacao}`, {
        method: 'GET', cfg, token,
      });
      if (getRes.status === 200) {
        const d = JSON.parse(getRes.body);
        // Considera pronto quando tiver linhaDigitavel ou status diferente de processando
        if (d.linhaDigitavel || (d.situacao && d.situacao !== 'EM_PROCESSAMENTO')) {
          cobranca = d;
          break;
        }
      }
    }

    if (!cobranca) {
      // Retorna o codigoSolicitacao para que o frontend possa tentar depois
      return res.json({
        codigoSolicitacao,
        linhaDigitavel: '',
        codigoBarras:   '',
        pdfLink:        '',
        pendente:       true,
      });
    }

    // Passo 3: busca PDF
    let pdfLink = '';
    const pdfRes = await interRequest(`/cobranca/v3/cobrancas/${codigoSolicitacao}/pdf`, {
      method: 'GET', cfg, token,
    });
    if (pdfRes.status === 200) {
      try {
        const pdfData = JSON.parse(pdfRes.body);
        if (pdfData.pdf) pdfLink = `data:application/pdf;base64,${pdfData.pdf}`;
      } catch (_) {}
    }

    res.json({
      codigoSolicitacao,
      nossoNumero:    cobranca.nossoNumero    || codigoSolicitacao,
      linhaDigitavel: cobranca.linhaDigitavel || '',
      codigoBarras:   cobranca.codigoBarras   || '',
      pdfLink,
      txid:           cobranca.txid           || '',
      pendente:       false,
    });

  } catch (err) {
    console.error('Erro ao gerar boleto:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Inter proxy (API V3) na porta ${PORT}`);
  console.log('Contas configuradas:', contasValidas.map(([n]) => n).join(', '));
});
