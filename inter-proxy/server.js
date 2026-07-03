/**
 * Rasteirissima — Proxy Banco Inter Cobranças
 *
 * Intermediário entre o frontend (index.html) e a API do Banco Inter.
 * Necessário porque a API do Inter exige mTLS (certificado digital),
 * que não pode ser usado diretamente pelo navegador.
 *
 * COMO USAR (desenvolvimento local):
 *   1. Copie cert.pem e key.pem (baixados do Inter Developers) para esta pasta
 *   2. Copie .env.example para .env e preencha as variáveis
 *   3. npm install
 *   4. node server.js
 *
 * DEPLOY (Railway ou Render — plano gratuito):
 *   - Suba esta pasta como repositório GitHub
 *   - Configure as variáveis de ambiente no painel do Railway/Render
 *   - O deploy é automático a cada push
 */

const express = require('express');
const https   = require('https');
const fs      = require('fs');

// ── Variáveis de ambiente ─────────────────────────────────────────────────────
const {
  INTER_CLIENT_ID,
  INTER_CLIENT_SECRET,
  INTER_CERT_PEM,        // conteúdo do cert.pem (string com \n)
  INTER_KEY_PEM,         // conteúdo do key.pem (string com \n)
  INTER_CONTA,           // número da conta Inter PJ (opcional, para alguns endpoints)
  ALLOWED_ORIGIN,        // ex: https://seuusuario.github.io
  PORT = 3000,
} = process.env;

// ── Validação mínima ──────────────────────────────────────────────────────────
const missing = ['INTER_CLIENT_ID','INTER_CLIENT_SECRET','INTER_CERT_PEM','INTER_KEY_PEM','ALLOWED_ORIGIN']
  .filter(k => !process.env[k]);
if (missing.length) {
  console.error('Variáveis de ambiente faltando:', missing.join(', '));
  process.exit(1);
}

// ── Agente HTTPS com mTLS (certificado digital do Inter) ─────────────────────
const interAgent = new https.Agent({
  cert: INTER_CERT_PEM.replace(/\\n/g, '\n'),
  key:  INTER_KEY_PEM.replace(/\\n/g, '\n'),
});

const INTER_BASE = 'https://cdpj.partners.bancointer.com.br';

// ── Cache de token OAuth2 ─────────────────────────────────────────────────────
let _token = null;
let _tokenExp = 0;

async function getToken() {
  if (_token && Date.now() < _tokenExp) return _token;

  const creds = Buffer.from(`${INTER_CLIENT_ID}:${INTER_CLIENT_SECRET}`).toString('base64');
  const res = await fetch(`${INTER_BASE}/oauth/v2/token`, {
    method: 'POST',
    headers: {
      'Authorization': `Basic ${creds}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: 'grant_type=client_credentials&scope=boleto-cobranca.read+boleto-cobranca.write',
    // @ts-ignore — Node 18+ suporta o agent no fetch global via dispatcher,
    // mas para compatibilidade usamos node-fetch se necessário
  });

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Inter OAuth falhou (${res.status}): ${txt}`);
  }
  const json = await res.json();
  _token = json.access_token;
  _tokenExp = Date.now() + (json.expires_in - 60) * 1000; // margem de 60s
  return _token;
}

// Wrapper de fetch com mTLS — Node 18+ com fetch nativo não suporta agents diretamente.
// Se precisar de versão anterior do Node, substituir por node-fetch ou axios.
async function interFetch(path, options = {}) {
  // Usa https.request manualmente para compatibilidade com mTLS
  return new Promise((resolve, reject) => {
    const url = new URL(INTER_BASE + path);
    const body = options.body ? Buffer.from(options.body, 'utf8') : null;
    const req = https.request(
      {
        hostname: url.hostname,
        path: url.pathname + url.search,
        method: options.method || 'GET',
        headers: {
          'Authorization': `Bearer ${options.token}`,
          'Content-Type': 'application/json',
          ...(body ? { 'Content-Length': body.length } : {}),
          ...(INTER_CONTA ? { 'x-conta-corrente': INTER_CONTA } : {}),
        },
        cert: INTER_CERT_PEM.replace(/\\n/g, '\n'),
        key:  INTER_KEY_PEM.replace(/\\n/g, '\n'),
      },
      (res) => {
        let data = '';
        res.on('data', c => { data += c; });
        res.on('end', () => resolve({ status: res.statusCode, body: data }));
      }
    );
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

// ── Servidor Express ──────────────────────────────────────────────────────────
const app = express();
app.use(express.json());

// CORS restrito ao domínio do sistema
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (origin === ALLOWED_ORIGIN || !origin) {
    res.setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN || '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  if (req.method === 'OPTIONS') { res.sendStatus(204); return; }
  next();
});

// ── GET /ping — verifica se o servidor está no ar ─────────────────────────────
app.get('/ping', (_req, res) => res.json({ ok: true }));

// ── POST /boleto — cria cobrança (boleto híbrido Boleto + Pix) ───────────────
//
// Body esperado do frontend:
// {
//   sacado: {
//     cnpjCpf, nome, email, cep, endereco, numero, complemento, bairro, cidade, uf, telefone
//   },
//   valor: 1234.56,           // número, em reais
//   vencimento: "2025-08-30", // ISO AAAA-MM-DD
//   descricao: "Boleto Marketing Jul/2025"  // texto livre (aparece no boleto)
// }
//
// Resposta para o frontend:
// {
//   nossoNumero, linhaDigitavel, codigoBarras, pdfLink, txid
// }
app.post('/boleto', async (req, res) => {
  try {
    const { sacado, valor, vencimento, descricao } = req.body;

    if (!sacado || !sacado.cnpjCpf || !valor || !vencimento) {
      return res.status(400).json({ error: 'Campos obrigatórios: sacado.cnpjCpf, valor, vencimento' });
    }

    const token = await getToken();

    // Payload para a API do Inter (Boleto Híbrido v3)
    // Documentação: https://developers.bancointer.com.br/reference/incluirboleto
    const payload = JSON.stringify({
      pagador: {
        cpfCnpj:     sacado.cnpjCpf.replace(/\D/g, ''),
        tipoPessoa:  sacado.cnpjCpf.replace(/\D/g, '').length === 11 ? 'FISICA' : 'juridica',
        nome:        sacado.nome,
        email:       sacado.email || '',
        telefone:    (sacado.telefone || '').replace(/\D/g, ''),
        endereco:    sacado.endereco || '',
        numero:      sacado.numero || 'S/N',
        complemento: sacado.complemento || '',
        bairro:      sacado.bairro || '',
        cidade:      sacado.cidade || '',
        uf:          (sacado.uf || '').toUpperCase(),
        cep:         sacado.cep.replace(/\D/g, ''),
      },
      valorNominal: Number(valor).toFixed(2),
      dataVencimento: vencimento,   // AAAA-MM-DD
      numDiasAgenda: 60,            // cancela automaticamente após 60 dias do vencimento
      mensagem: {
        linha1: (descricao || 'Rasteirissima Franchising').slice(0, 50),
      },
    });

    const interRes = await interFetch('/cobranca/v3/boletos', {
      method: 'POST',
      token,
      body: payload,
    });

    if (interRes.status !== 200 && interRes.status !== 201) {
      return res.status(502).json({ error: `Inter respondeu ${interRes.status}`, detail: interRes.body });
    }

    const data = JSON.parse(interRes.body);

    // Busca PDF do boleto (GET /cobranca/v3/boletos/{nossoNumero}/pdf)
    let pdfLink = '';
    if (data.nossoNumero) {
      const pdfRes = await interFetch(`/cobranca/v3/boletos/${data.nossoNumero}/pdf`, {
        method: 'GET',
        token,
      });
      if (pdfRes.status === 200) {
        // Inter retorna { pdf: "<base64>" }
        try {
          const pdfData = JSON.parse(pdfRes.body);
          if (pdfData.pdf) {
            // Devolve como data URI para o frontend abrir direto
            pdfLink = `data:application/pdf;base64,${pdfData.pdf}`;
          }
        } catch (_) {}
      }
    }

    res.json({
      nossoNumero:   data.nossoNumero   || '',
      linhaDigitavel: data.linhaDigitavel || '',
      codigoBarras:  data.codigoBarras  || '',
      pdfLink,
      txid:          data.txid          || '',
    });
  } catch (err) {
    console.error('Erro ao gerar boleto:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Inter proxy rodando na porta ${PORT}`));
