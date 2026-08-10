const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const http = require('http');
require('dotenv').config();

const PORT = Number(process.env.REDEEM_API_PORT || 8787);
const BOT_TOKEN = process.env.BOT_TOKEN || '';
const VOUCHERS_FILE = process.env.VOUCHERS_FILE || path.join(__dirname, 'data', 'vouchers.json');
const USERS_FILE = process.env.USERS_FILE || path.join(__dirname, 'data', 'users.json');
const ENTITLEMENTS_FILE = process.env.ENTITLEMENTS_FILE || path.join(__dirname, 'data', 'entitlements.json');

const RATE_LIMIT_MAX = Number(process.env.REDEEM_RATE_LIMIT_MAX || 5);
const RATE_LIMIT_WINDOW_MS = Number(process.env.REDEEM_RATE_LIMIT_WINDOW_MS || 10 * 60 * 1000);
const INIT_DATA_MAX_AGE_SECONDS = Number(process.env.INIT_DATA_MAX_AGE_SECONDS || 24 * 60 * 60);

async function ensureJsonFile(filePath, initial) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  try {
    await fs.access(filePath);
  } catch {
    await fs.writeFile(filePath, JSON.stringify(initial, null, 2));
  }
}

async function ensureVoucherFile() {
  await ensureJsonFile(VOUCHERS_FILE, { packages: {}, issued: {} });
}

async function ensureUsersFile() {
  await ensureJsonFile(USERS_FILE, { users: {}, requests: {}, drafts: {} });
}

async function ensureEntitlementsFile() {
  await ensureJsonFile(ENTITLEMENTS_FILE, { entitlements: {} });
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

async function readUserStore() {
  await ensureUsersFile();
  const raw = await fs.readFile(USERS_FILE, 'utf8');
  const store = JSON.parse(raw);
  return {
    users: store.users || {},
    requests: store.requests || {},
    drafts: store.drafts || {}
  };
}

async function readEntitlements() {
  await ensureEntitlementsFile();
  const raw = await fs.readFile(ENTITLEMENTS_FILE, 'utf8');
  const store = JSON.parse(raw);
  return {
    entitlements: store.entitlements || {}
  };
}

async function writeEntitlements(store) {
  await ensureEntitlementsFile();
  await fs.writeFile(ENTITLEMENTS_FILE, JSON.stringify({
    entitlements: store.entitlements || {}
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

function verifyInitData(initData) {
  if (!BOT_TOKEN || !initData) {
    return null;
  }

  const params = new URLSearchParams(String(initData));
  const hash = params.get('hash');
  if (!hash) {
    return null;
  }

  params.delete('hash');

  const dataCheckString = [...params.keys()]
    .sort()
    .map((key) => `${key}=${params.get(key)}`)
    .join('\n');

  const secretKey = crypto.createHmac('sha256', 'WebAppData').update(BOT_TOKEN).digest();
  const computedHash = crypto.createHmac('sha256', secretKey).update(dataCheckString).digest('hex');

  if (computedHash !== hash) {
    return null;
  }

  const authDate = Number(params.get('auth_date') || 0);
  if (!authDate || (Date.now() / 1000) - authDate > INIT_DATA_MAX_AGE_SECONDS) {
    return null;
  }

  let user = null;
  try {
    const rawUser = params.get('user');
    if (rawUser) {
      user = JSON.parse(rawUser);
    }
  } catch {
    user = null;
  }

  if (!user || !user.id) {
    return null;
  }

  return {
    id: String(user.id),
    username: user.username || '',
    firstName: user.first_name || '',
    lastName: user.last_name || ''
  };
}

const redeemAttempts = new Map();

function isRateLimited(key) {
  const now = Date.now();
  const cutoff = now - RATE_LIMIT_WINDOW_MS;
  const recent = (redeemAttempts.get(key) || []).filter((timestamp) => timestamp > cutoff);
  recent.push(now);
  redeemAttempts.set(key, recent);
  return recent.length > RATE_LIMIT_MAX;
}

function findVoucher(store, phrase) {
  for (const pool of Object.values(store.packages || {})) {
    for (const voucher of Object.values(pool.issued || {})) {
      if (normalizePhrase(voucher.phrase) === phrase) {
        return voucher;
      }
    }
  }
  return null;
}

let redemptionQueue = Promise.resolve();

// Serializes read-modify-write redemption operations so two simultaneous
// requests can never both redeem the same voucher (ASSIGNED -> REDEEMED once).
function runExclusiveRedemption(task) {
  const result = redemptionQueue.then(task, task);
  redemptionQueue = result.then(() => undefined, () => undefined);
  return result;
}

async function redeemPhrase({ code, phrase, initData, deviceId, ip }) {
  const remoteKey = `ip:${ip || 'unknown'}`;
  if (isRateLimited(remoteKey)) {
    return { status: 429, body: { ok: false, error: 'RATE_LIMITED' } };
  }

  const user = verifyInitData(initData);
  if (!user) {
    return { status: 401, body: { ok: false, error: 'UNAUTHORIZED' } };
  }

  if (isRateLimited(`user:${user.id}`)) {
    return { status: 429, body: { ok: false, error: 'RATE_LIMITED' } };
  }

  const voucherCode = normalizePhrase(code || phrase || '');
  if (!voucherCode) {
    return { status: 400, body: { ok: false, error: 'BAD_REQUEST' } };
  }

  const result = await runExclusiveRedemption(async () => {
    const voucherStore = await readVoucherStore();
    const voucher = findVoucher(voucherStore, voucherCode);

    if (!voucher) {
      return { ok: false, error: 'INVALID_CODE' };
    }

    if (voucher.status === 'revoked') {
      return { ok: false, error: 'VOUCHER_REVOKED' };
    }

    if (voucher.status === 'redeemed') {
      return { ok: false, error: 'ALREADY_REDEEMED' };
    }

    if (voucher.status !== 'assigned' && voucher.status !== 'reserved') {
      return { ok: false, error: 'NOT_ASSIGNED' };
    }

    if (String(voucher.assignedToUserId || voucher.userId) !== user.id) {
      return { ok: false, error: 'NOT_OWNER' };
    }

    const userStore = await readUserStore();
    const request = (userStore.requests || {})[String(voucher.requestId)] || null;
    if (!request || request.status !== 'approved') {
      return { ok: false, error: 'PAYMENT_NOT_APPROVED' };
    }

    const redeemedAt = new Date().toISOString();
    voucher.status = 'redeemed';
    voucher.redeemedAt = redeemedAt;
    voucher.redeemedByUserId = user.id;
    voucher.redeemedByUsername = user.username || '';
    voucher.redeemedByDeviceId = String(deviceId || '');

    const poolKey = voucher.packagePool;
    if (poolKey && voucherStore.packages[poolKey]) {
      voucherStore.packages[poolKey].issued[String(voucher.requestId)] = voucher;
    }
    voucherStore.issued[String(voucher.requestId)] = voucher;
    await writeVoucherStore(voucherStore);

    const entitlements = await readEntitlements();
    entitlements.entitlements[user.id] = {
      packageKey: voucher.packageKey,
      packageLabel: voucher.packageLabel || voucher.packageKey,
      requestId: voucher.requestId,
      redeemedAt,
      redeemedByDeviceId: String(deviceId || '')
    };
    await writeEntitlements(entitlements);

    return {
      ok: true,
      packageKey: voucher.packageKey,
      packageLabel: voucher.packageLabel || voucher.packageKey,
      redeemedAt,
      access: {
        packageKey: voucher.packageKey,
        enabled: true
      }
    };
  });

  if (result.ok) {
    return { status: 200, body: result };
  }

  return { status: 400, body: result };
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
    const forwarded = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim();
    const ip = forwarded || (req.socket && req.socket.remoteAddress) || '';
    const result = await redeemPhrase({ ...body, ip });
    sendJson(res, result.status, result.body);
  } catch (error) {
    console.error('Redemption error:', error);
    sendJson(res, 400, { ok: false, error: 'BAD_REQUEST' });
  }
});

server.listen(PORT, () => {
  console.log(`Voucher redemption API listening on http://127.0.0.1:${PORT}`);
});
