/**
 * Rasteirissima — Proxy Banco Inter Cobranças
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
// Cada conta precisa de: CLIENT_ID, CLIENT_SECRET, CERT_PEM, KEY_PEM
// Os nomes das env vars seguem o padrão INTER_<CONTA>_<CAMPO>
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

// Valida que pelo menos uma conta tem credenciais
const contasValidas = Object.entries(CONTAS).filter(([, c]) => c.clientId && c.certPem);
if (!contasValidas.length) {
  console.error('Nenhuma conta Inter configurada. Defina INTER_BRI_CLIENT_ID, INTER_MKT_CLIENT_ID ou INTER_FRAN_CLIENT_ID nas variáveis de ambiente.');
  process.exit(1);
}
if (!ALLOWED_ORIGIN) {
  console.warn('ALLOWED_ORIGIN não definido — CORS aberto (não recomendado em produção)');
}

const INTER_BASE = 'https://cdpj.partners.bancointer.com.br';

// ── OAuth2: token por conta com cache ────────────────────────────────────────
async function getToken(cfg) {
  if (cfg.token && Date.now() < cfg.tokenExp) return cfg.token;

  const creds = Buffer.from(`${cfg.clientId}:${cfg.clientSecret}`).toString('base64');
  const body = 'grant_type=client_credentials&scope=boleto-cobranca.read+boleto-cobranca.write';

  const result = await interRequest('/oauth/v2/token', {
    method: 'POST',
    cfg,
    headers: {
      'Authorization': `Basic ${creds}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
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

    const req = https.request(
      {
        hostname: url.hostname,
        path: url.pathname + url.search,
        method,
        headers: reqHeaders,
        cert: cfg.certPem,
        key:  cfg.keyPem,
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

// GET /ping — verifica saúde do servidor e quais contas estão configuradas
app.get('/ping', (_req, res) => {
  res.json({
    ok: true,
    contas: Object.entries(CONTAS)
      .filter(([, c]) => c.clientId)
      .map(([nome]) => nome),
  });
});

// POST /boleto — cria cobrança (boleto híbrido Boleto + Pix)
//
// Body esperado:
// {
//   conta: 'brilhante' | 'marketing' | 'franchising',
//   sacado: { cnpjCpf, nome, email, cep, endereco, numero, complemento, bairro, cidade, uf, telefone },
//   valor: 1234.56,
//   vencimento: 'AAAA-MM-DD',
//   descricao: 'Boleto Marketing — NF 123 — Loja XYZ'
// }
app.post('/boleto', async (req, res) => {
  try {
    const { conta, sacado, valor, vencimento, descricao } = req.body;

    if (!conta || !CONTAS[conta]) {
      return res.status(400).json({ error: `Conta inválida: "${conta}". Use brilhante, marketing ou franchising.` });
    }
    if (!sacado || !sacado.cnpjCpf || !valor || !vencimento) {
      return res.status(400).json({ error: 'Campos obrigatórios: sacado.cnpjCpf, valor, vencimento' });
    }

    const cfg = CONTAS[conta];
    if (!cfg.clientId || !cfg.certPem) {
      return res.status(503).json({ error: `Conta "${conta}" não está configurada neste servidor.` });
    }

    const token = await getToken(cfg);

    const cnpjCpfLimpo = sacado.cnpjCpf.replace(/\D/g, '');
    const payload = JSON.stringify({
      pagador: {
        cpfCnpj:     cnpjCpfLimpo,
        tipoPessoa:  cnpjCpfLimpo.length === 11 ? 'FISICA' : 'JURIDICA',
        nome:        sacado.nome,
        email:       sacado.email     || '',
        telefone:    (sacado.telefone || '').replace(/\D/g, ''),
        endereco:    sacado.endereco  || '',
        numero:      sacado.numero    || 'S/N',
        complemento: sacado.complemento || '',
        bairro:      sacado.bairro    || '',
        cidade:      sacado.cidade    || '',
        uf:          (sacado.uf || '').toUpperCase().slice(0, 2),
        cep:         sacado.cep.replace(/\D/g, ''),
      },
      valorNominal:   Number(valor).toFixed(2),
      dataVencimento: vencimento,
      numDiasAgenda:  60,
      mensagem: {
        linha1: (descricao || 'Rasteirissima').slice(0, 50),
      },
    });

    const createRes = await interRequest('/cobranca/v3/boletos', {
      method: 'POST',
      cfg,
      token,
      headers: { 'Content-Type': 'application/json' },
      body: payload,
    });

    if (createRes.status !== 200 && createRes.status !== 201) {
      return res.status(502).json({
        error: `Inter respondeu ${createRes.status}`,
        detail: createRes.body,
      });
    }

    const data = JSON.parse(createRes.body);

    // Busca PDF do boleto (base64)
    let pdfLink = '';
    if (data.nossoNumero) {
      const pdfRes = await interRequest(`/cobranca/v3/boletos/${data.nossoNumero}/pdf`, {
        method: 'GET', cfg, token,
      });
      if (pdfRes.status === 200) {
        try {
          const pdfData = JSON.parse(pdfRes.body);
          if (pdfData.pdf) pdfLink = `data:application/pdf;base64,${pdfData.pdf}`;
        } catch (_) {}
      }
    }

    res.json({
      nossoNumero:    data.nossoNumero    || '',
      linhaDigitavel: data.linhaDigitavel || '',
      codigoBarras:   data.codigoBarras   || '',
      pdfLink,
      txid:           data.txid           || '',
    });
  } catch (err) {
    console.error('Erro ao gerar boleto:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Inter proxy na porta ${PORT}`);
  console.log('Contas configuradas:', contasValidas.map(([n]) => n).join(', '));
});
