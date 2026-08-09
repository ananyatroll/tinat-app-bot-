const fs = require('fs/promises');
const path = require('path');
require('dotenv').config();
const http = require('http');

const PORT = Number(process.env.REDEEM_API_PORT || 8787);
const VOUCHERS_FILE = process.env.VOUCHERS_FILE || path.join(__dirname, 'data', 'vouchers.json');

async function ensureVoucherFile() {
  await fs.mkdir(path.dirname(VOUCHERS_FILE), { recursive: true });
  try {
    await fs.access(VOUCHERS_FILE);
  } catch {
    await fs.writeFile(VOUCHERS_FILE, JSON.stringify({ packages: {}, issued: {} }, null, 2));
  }
}

async function readVoucherStore() {
  await ensureVoucherFile();
  const raw = await fs.readFile(VOUCHERS_FILE, 'utf8');
  const store = JSON.parse(raw);
  return {
    packages: store.packages || {},
    issued: store.issued || {}
  };
}

async function writeVoucherStore(store) {
  await ensureVoucherFile();
  await fs.writeFile(VOUCHERS_FILE, JSON.stringify({
    packages: store.packages || {},
    issued: store.issued || {}
  }, null, 2));
}

function normalizePhrase(phrase) {
  return String(phrase || '').trim();
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { 'content-type': 'application/json' });
  res.end(JSON.stringify(payload));
}

async function parseJsonBody(req) {
  let body = '';
  for await (const chunk of req) {
    body += chunk;
  }

  if (!body.trim()) {
    return {};
  }

  return JSON.parse(body);
}

async function redeemPhrase({ phrase, userId, deviceId }) {
  const store = await readVoucherStore();
  const normalizedPhrase = normalizePhrase(phrase);

  let foundVoucher = null;
  let foundPoolKey = null;

  for (const [poolKey, pool] of Object.entries(store.packages || {})) {
    const issuedMap = pool.issued || {};
    for (const voucher of Object.values(issuedMap)) {
      if (normalizePhrase(voucher.phrase) === normalizedPhrase) {
        foundVoucher = voucher;
        foundPoolKey = poolKey;
        break;
      }
    }
    if (foundVoucher) {
      break;
    }
  }

  if (!foundVoucher) {
    return { ok: false, error: 'INVALID_PHRASE' };
  }

  if (foundVoucher.status === 'redeemed') {
    return { ok: false, error: 'ALREADY_REDEEMED' };
  }

  const redeemedAt = new Date().toISOString();
  foundVoucher.status = 'redeemed';
  foundVoucher.redeemedAt = redeemedAt;
  foundVoucher.redeemedByUserId = String(userId || '');
  foundVoucher.redeemedByDeviceId = String(deviceId || '');

  store.packages[foundPoolKey].issued[foundVoucher.requestId] = foundVoucher;
  store.issued[foundVoucher.requestId] = foundVoucher;
  await writeVoucherStore(store);

  return {
    ok: true,
    packageKey: foundVoucher.packageKey,
    packageLabel: foundVoucher.packageLabel || foundVoucher.packageKey,
    redeemedAt,
    access: {
      packageKey: foundVoucher.packageKey,
      enabled: true
    }
  };
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'GET' && req.url === '/health') {
    sendJson(res, 200, { ok: true });
    return;
  }

  if (req.method !== 'POST' || req.url !== '/api/v1/vouchers/redeem') {
    sendJson(res, 404, { ok: false, error: 'NOT_FOUND' });
    return;
  }

  try {
    const body = await parseJsonBody(req);
    const result = await redeemPhrase(body);
    sendJson(res, result.ok ? 200 : 400, result);
  } catch (error) {
    sendJson(res, 400, { ok: false, error: 'BAD_REQUEST' });
  }
});

server.listen(PORT, () => {
  console.log(`Voucher redemption API listening on http://127.0.0.1:${PORT}`);
});