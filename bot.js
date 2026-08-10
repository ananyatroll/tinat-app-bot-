const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
require('dotenv').config();
const { Telegraf } = require('telegraf');
const ExcelJS = require('exceljs');

const BOT_TOKEN = process.env.BOT_TOKEN;
const ADMIN_CHAT_ID = process.env.ADMIN_CHAT_ID;
const PHRASES_DIR = process.env.PHRASES_DIR || path.join(__dirname, 'phrases');
const VOUCHERS_FILE = process.env.VOUCHERS_FILE || path.join(__dirname, 'data', 'vouchers.json');
const EXPORTS_DIR = process.env.EXPORTS_DIR || path.join(__dirname, 'exports');
const PACKAGES = loadPackages();
const DATA_FILE = path.join(__dirname, 'data', 'users.json');

if (!BOT_TOKEN) {
  throw new Error('Missing BOT_TOKEN in environment');
}

const bot = new Telegraf(BOT_TOKEN);

async function ensureDataFile() {
  await fs.mkdir(path.dirname(DATA_FILE), { recursive: true });
  try {
    await fs.access(DATA_FILE);
  } catch {
    await fs.writeFile(DATA_FILE, JSON.stringify({ users: {}, requests: {}, drafts: {} }, null, 2));
  }
}

async function ensureVoucherFile() {
  await fs.mkdir(path.dirname(VOUCHERS_FILE), { recursive: true });
  try {
    await fs.access(VOUCHERS_FILE);
  } catch {
    await fs.writeFile(VOUCHERS_FILE, JSON.stringify({ packages: {}, issued: {} }, null, 2));
  }
}

async function ensureExportsDir() {
  await fs.mkdir(EXPORTS_DIR, { recursive: true });
}

async function readStore() {
  await ensureDataFile();
  const raw = await fs.readFile(DATA_FILE, 'utf8');
  const store = JSON.parse(raw);
  return {
    users: store.users || {},
    requests: store.requests || {},
    drafts: store.drafts || {}
  };
}

async function writeStore(store) {
  await ensureDataFile();
  await fs.writeFile(DATA_FILE, JSON.stringify({
    users: store.users || {},
    requests: store.requests || {},
    drafts: store.drafts || {}
  }, null, 2));
}

function randomString(length) {
  return crypto.randomBytes(length).toString('base64url').slice(0, length);
}

function formatMoney(cents, currency = 'USD') {
  return `${(cents / 100).toFixed(2)} ${currency}`;
}

function loadPackages() {
  const rawPackages = process.env.PACKAGES_JSON;
  if (!rawPackages) {
    return [{
      key: 'default',
      label: process.env.ACCESS_LABEL || 'Tinat Access',
      priceCents: Number(process.env.PRICE_CENTS || 5000),
      currency: (process.env.CURRENCY || 'USD').toUpperCase(),
      phrasePool: 'default'
    }];
  }

  let parsed;
  try {
    parsed = JSON.parse(rawPackages);
  } catch {
    throw new Error('PACKAGES_JSON must be valid JSON');
  }

  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error('PACKAGES_JSON must contain at least one package');
  }

  return parsed.map((pkg, index) => ({
    key: String(pkg.key || pkg.id || `package_${index + 1}`).trim(),
    label: String(pkg.label || pkg.name || `Package ${index + 1}`).trim(),
    priceCents: Number(pkg.priceCents || pkg.price_cents || 0),
    currency: String(pkg.currency || process.env.CURRENCY || 'USD').toUpperCase(),
    phrasePool: String(pkg.phrasePool || pkg.pool || pkg.key || pkg.id || `package_${index + 1}`).trim()
  })).filter((pkg) => pkg.key && pkg.label && Number.isFinite(pkg.priceCents) && pkg.priceCents > 0);
}

function getPackageList() {
  return PACKAGES;
}

function getPackageByKey(packageKey) {
  return PACKAGES.find((pkg) => pkg.key === packageKey) || null;
}

function buildPackageKeyboard(packages) {
  return {
    inline_keyboard: packages.map((pkg) => ([{
      text: `${pkg.label} - ${formatMoney(pkg.priceCents, pkg.currency)}`,
      callback_data: `package:${pkg.key}`
    }]))
  };
}

const PAYMENT_METHODS = [
  { key: 'cbe', label: 'Commercial Bank of Ethiopia (CBE)' },
  { key: 'telebirr', label: 'Telebirr' }
];

const PAYMENT_LINK_PATTERNS = {
  cbe: /^https:\/\/mbreciept\.cbe\.com\.et\/\S+/i,
  telebirr: /^https:\/\/transactioninfo\.ethiotelecom\.et\/receipt\/\S+/i
};

const PAYMENT_ID_PATTERNS = {
  cbe: /^FT\d{4,}[A-Z0-9]{3,}$/i,
  telebirr: /^[A-Z]{3}\d{1,}[A-Z0-9]{2,}$/i
};

function isValidTransactionLink(paymentMethodKey, link) {
  const pattern = PAYMENT_LINK_PATTERNS[paymentMethodKey];
  return pattern ? pattern.test(normalizeTransactionLink(link)) : false;
}

function isValidTransactionId(paymentMethodKey, value) {
  const pattern = PAYMENT_ID_PATTERNS[paymentMethodKey];
  return pattern ? pattern.test(normalizeTransactionId(value)) : false;
}

function getLinkFormatHint(paymentMethodKey) {
  if (paymentMethodKey === 'cbe') {
    return 'https://mbreciept.cbe.com.et/<your-receipt-code>';
  }
  if (paymentMethodKey === 'telebirr') {
    return 'https://transactioninfo.ethiotelecom.et/receipt/<your-transaction-id>';
  }
  return 'the full receipt link';
}

function getIdFormatHint(paymentMethodKey) {
  if (paymentMethodKey === 'cbe') {
    return 'starts with FT, e.g. FT26222QKMBG';
  }
  if (paymentMethodKey === 'telebirr') {
    return 'letters and digits, e.g. DGJ22CMPJM';
  }
  return 'your transaction ID';
}

function getPaymentMethodByKey(methodKey) {
  return PAYMENT_METHODS.find((method) => method.key === methodKey) || null;
}

function buildPaymentMethodKeyboard() {
  return {
    inline_keyboard: PAYMENT_METHODS.map((method) => ([{
      text: method.label,
      callback_data: `method:${method.key}`
    }]))
  };
}

function normalizePhrase(phrase) {
  return String(phrase || '').trim();
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

function formatDateForSheet(value) {
  return value ? new Date(value).toISOString() : '';
}

async function createApprovedWorkbook(store) {
  await ensureExportsDir();

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'tinat-bot';
  workbook.created = new Date();

  // One sheet per configured package, plus a catch-all for unknown labels.
  const workbookData = {};
  for (const pkg of getPackageList()) {
    workbookData[pkg.label] = [];
  }
  workbookData['Other'] = [];

  const approvedRequests = Object.values(store.requests || {})
    .filter((req) => req.status === 'approved')
    .sort((a, b) => String(a.reviewedAt || '').localeCompare(String(b.reviewedAt || '')));

  for (const req of approvedRequests) {
    const user = req.user || {};
    const sheetName = workbookData[req.packageLabel] ? req.packageLabel : 'Other';
    workbookData[sheetName].push({
      requestId: req.requestId,
      packageLabel: req.packageLabel,
      price: formatMoney(req.priceCents, req.currency),
      userName: `${user.firstName || ''} ${user.lastName || ''}`.trim(),
      telegramUsername: user.username || '',
      telegramId: user.id,
    transactionId: req.transactionId,
    transactionLink: req.transactionLink,
    paymentMethod: req.paymentMethodLabel || req.paymentMethod || '',
      voucherPhrase: req.voucherPhrase || '',
      status: req.status,
      approvedAt: formatDateForSheet(req.reviewedAt || new Date().toISOString()),
      approvedBy: String(req.reviewedBy || '')
    });
  }

  for (const [currentSheetName, rows] of Object.entries(workbookData)) {
    if (currentSheetName === 'Other' && rows.length === 0) {
      continue;
    }

    const sheet = workbook.addWorksheet(currentSheetName);
    sheet.columns = [
      { header: 'Request ID', key: 'requestId', width: 20 },
      { header: 'Package', key: 'packageLabel', width: 24 },
      { header: 'Price', key: 'price', width: 14 },
      { header: 'Payment Method', key: 'paymentMethod', width: 22 },
      { header: 'User Name', key: 'userName', width: 22 },
      { header: 'Telegram Username', key: 'telegramUsername', width: 20 },
      { header: 'Telegram ID', key: 'telegramId', width: 16 },
      { header: 'Transaction ID', key: 'transactionId', width: 20 },
      { header: 'Transaction Link', key: 'transactionLink', width: 40 },
      { header: 'Voucher Phrase', key: 'voucherPhrase', width: 32 },
      { header: 'Status', key: 'status', width: 14 },
      { header: 'Approved At', key: 'approvedAt', width: 24 },
      { header: 'Approved By', key: 'approvedBy', width: 16 }
    ];

    for (const row of rows) {
      sheet.addRow(row);
    }

    sheet.getRow(1).font = { bold: true };
    sheet.autoFilter = `A1:M${Math.max(rows.length + 1, 1)}`;
  }

  const fileName = `approved-${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
  const filePath = path.join(EXPORTS_DIR, fileName);
  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

function countApprovedRequests(store) {
  return Object.values(store.requests || {})
    .filter((req) => req.status === 'approved')
    .length;
}

async function exportApprovedExcel(chatId) {
  const store = await readStore();
  const count = countApprovedRequests(store);
  const workbookPath = await createApprovedWorkbook(store);
  await bot.telegram.sendDocument(chatId, { source: workbookPath, filename: path.basename(workbookPath) }, {
    caption: `Full export of ${count} approved transaction(s).`
  });
}

async function hydrateVoucherStoreFromPhraseFiles(store) {
  let changed = false;

  const globalIssued = new Set(
    Object.values(store.issued || {})
      .map((record) => normalizePhrase(record.phrase))
      .filter(Boolean)
  );

  for (const pkg of getPackageList()) {
    const pool = store.packages[pkg.phrasePool] || { available: [], issued: {} };
    pool.issued = pool.issued || {};

    if (Array.isArray(pool.available) && pool.available.length > 0) {
      continue;
    }

    const sourceFile = path.join(PHRASES_DIR, `${pkg.phrasePool}.txt`);
    try {
      const raw = await fs.readFile(sourceFile, 'utf8');
      const phrases = raw
        .split(/\r?\n/)
        .map(normalizePhrase)
        .filter(Boolean);

      // Never re-add phrases that have already been handed out to anyone,
      // so a code can never be issued to two different people.
      const freshPhrases = phrases.filter((phrase) => !globalIssued.has(phrase));

      pool.available = freshPhrases;
      store.packages[pkg.phrasePool] = pool;
      changed = true;
    } catch {
      store.packages[pkg.phrasePool] = pool;
    }
  }

  if (changed) {
    await writeVoucherStore(store);
  }

  return store;
}

async function reserveVoucherPhrase(packageKey, request) {
  const pkg = getPackageByKey(packageKey);
  if (!pkg) {
    throw new Error(`Unknown package: ${packageKey}`);
  }

  const store = await hydrateVoucherStoreFromPhraseFiles(await readVoucherStore());
  const pool = store.packages[pkg.phrasePool] || { available: [], issued: {} };
  pool.issued = pool.issued || {};

  const globalIssued = new Set(
    Object.values(store.issued || {})
      .map((record) => normalizePhrase(record.phrase))
      .filter(Boolean)
  );

  let phrase = '';
  while (Array.isArray(pool.available) && pool.available.length > 0) {
    const candidate = normalizePhrase(pool.available.shift());
    if (candidate && !globalIssued.has(candidate)) {
      phrase = candidate;
      break;
    }
  }

  if (!phrase) {
    throw new Error(`No voucher phrases left for package ${pkg.label}`);
  }

  pool.issued[String(request.requestId)] = {
    requestId: request.requestId,
    userId: request.user.id,
    packageKey: pkg.key,
    packageLabel: pkg.label,
    packagePool: pkg.phrasePool,
    phrase,
    status: 'reserved',
    reservedAt: new Date().toISOString()
  };

  store.packages[pkg.phrasePool] = pool;
  store.issued[String(request.requestId)] = pool.issued[String(request.requestId)];
  await writeVoucherStore(store);
  return pool.issued[String(request.requestId)];
}

async function getVoucherForUser(userId) {
  const store = await readVoucherStore();
  const records = Object.values(store.issued || {})
    .filter((record) => String(record.userId) === String(userId) && normalizePhrase(record.phrase))
    .sort((a, b) => String(b.reservedAt || '').localeCompare(String(a.reservedAt || '')));
  return records[0] || null;
}

function buildAdminReviewMessage(request) {
  return [
    'New access request pending review:',
    `Request ID: ${request.requestId}`,
    `Package: ${request.packageLabel} (${formatMoney(request.priceCents, request.currency)})`,
    `User: ${request.user.firstName || ''} ${request.user.lastName || ''}`.trim(),
    `Telegram: @${request.user.username || 'no_username'} (${request.user.id})`,
    `Transaction ID: ${maskTransactionId(request.transactionId)}`,
    `Transaction Link: ${request.transactionLink}`,
    '',
    'Approve only after verifying the payment.'
  ].join('\n');
}

function normalizeTransactionLink(value) {
  return String(value || '').trim();
}

function normalizeTransactionId(value) {
  return String(value || '').trim();
}

function maskTransactionId(value) {
  const id = String(value || '').trim();
  if (id.length <= 2) {
    return '*'.repeat(id.length);
  }
  return id.slice(0, 2) + '*'.repeat(id.length - 2);
}

function getPackageSelectionMessage(packages) {
  return [
    'Choose your package first:',
    ...packages.map((pkg) => `${pkg.label} - ${formatMoney(pkg.priceCents, pkg.currency)}`),
    '',
    'I will then ask for your payment method, name, transaction link, and transaction ID.'
  ].join('\n');
}

function getStartMessage(packages) {
  return [
    'Welcome to Tinat.',
    'Pick a package below, then I will ask for your payment method, name, transaction link, and transaction ID.',
    '',
    ...packages.map((pkg) => `${pkg.label} - ${formatMoney(pkg.priceCents, pkg.currency)}`),
    '',
    'After approval, the bot will send you one secret phrase for your selected package.'
  ].join('\n');
}

function getDraft(store, userId) {
  return store.drafts[String(userId)] || null;
}

function setDraft(store, userId, draft) {
  store.drafts[String(userId)] = draft;
}

function clearDraft(store, userId) {
  delete store.drafts[String(userId)];
}

function getRequest(store, requestId) {
  return store.requests[String(requestId)] || null;
}

function setRequest(store, requestId, request) {
  store.requests[String(requestId)] = request;
}

function clearRequest(store, requestId) {
  delete store.requests[String(requestId)];
}

function buildRequestMessage(request) {
  return [
    'New access request pending review:',
    `Request ID: ${request.requestId}`,
    `Package: ${request.packageLabel} (${formatMoney(request.priceCents, request.currency)})`,
    `Payment Method: ${request.paymentMethodLabel || request.paymentMethod || 'N/A'}`,
    `User: ${request.user.firstName || ''} ${request.user.lastName || ''}`.trim(),
    `Telegram: @${request.user.username || 'no_username'} (${request.user.id})`,
    `Transaction ID: ${maskTransactionId(request.transactionId)}`,
    `Transaction Link: ${request.transactionLink}`,
    '',
    'Approve this request only after verifying the payment.'
  ].join('\n');
}

function buildPendingMessage(request) {
  return [
    'Your payment details were submitted.',
    `Request ID: ${request.requestId}`,
    'I sent it to the admin for verification.',
    'You will get your voucher phrase after approval.'
  ].join('\n');
}

function buildApprovedMessage(record) {
  return [
    'Payment approved. Your Tinat voucher phrase is below:',
    `Package: ${record.packageLabel}`,
    `Phrase: ${record.voucherPhrase}`,
    '',
    'Keep this phrase safe. If you lose it, send /myaccess in this bot.'
  ].join('\n');
}

async function notifyAdminAboutRequest(request) {
  if (!ADMIN_CHAT_ID) {
    throw new Error('Missing ADMIN_CHAT_ID in environment');
  }

  await bot.telegram.sendMessage(ADMIN_CHAT_ID, buildRequestMessage(request), {
    reply_markup: {
      inline_keyboard: [[
        { text: 'Approve', callback_data: `approve:${request.requestId}` },
        { text: 'Reject', callback_data: `reject:${request.requestId}` }
      ]]
    }
  });
}

let voucherMutationQueue = Promise.resolve();

// Serializes read-modify-write operations on the voucher store so two
// approvals can never reserve the same phrase (no double counting).
function runExclusiveVoucherMutation(task) {
  const result = voucherMutationQueue.then(task, task);
  voucherMutationQueue = result.then(() => undefined, () => undefined);
  return result;
}

async function finalizeRequest(requestId, approvedBy, approved) {
  const store = await readStore();
  const request = getRequest(store, requestId);

  if (!request) {
    return null;
  }

  if (request.status !== 'pending') {
    return request;
  }

  request.status = approved ? 'approved' : 'rejected';
  request.reviewedAt = new Date().toISOString();
  request.reviewedBy = approvedBy;
  setRequest(store, requestId, request);

  if (!approved) {
    clearDraft(store, request.user.id);
    await writeStore(store);
    return request;
  }

  const voucher = await runExclusiveVoucherMutation(() => reserveVoucherPhrase(request.packageKey, request));

  request.deliveredAt = new Date().toISOString();
  request.voucherPhrase = voucher.phrase;
  request.voucherStatus = 'delivered';
  request.excelExportAt = new Date().toISOString();
  clearDraft(store, request.user.id);
  clearRequest(store, requestId);
  setRequest(store, requestId, request);
  await writeStore(store);

  const workbookPath = await createApprovedWorkbook(store);

  await bot.telegram.sendMessage(request.user.id, buildApprovedMessage({
    packageLabel: request.packageLabel,
    voucherPhrase: voucher.phrase
  }));

  if (ADMIN_CHAT_ID) {
    const count = countApprovedRequests(store);
    await bot.telegram.sendDocument(ADMIN_CHAT_ID, { source: workbookPath, filename: path.basename(workbookPath) }, {
      caption: `Excel updated: ${count} approved transaction(s) (incl. request ${request.requestId}). Use /export to re-download anytime.`
    });
  }

  return request;
}

async function submitRequest(ctx, draft) {
  const store = await readStore();
  const existingAccess = store.users[String(ctx.from.id)];

  if (existingAccess) {
    await ctx.reply('You already have approved access. Use /myaccess to get your login details again.');
    clearDraft(store, ctx.from.id);
    await writeStore(store);
    return;
  }

  const requestId = randomString(10);
  const request = {
    requestId,
    status: 'pending',
    submittedAt: new Date().toISOString(),
    packageKey: draft.packageKey,
    packageLabel: draft.packageLabel,
    priceCents: draft.priceCents,
    currency: draft.currency,
    user: {
      id: ctx.from.id,
      username: ctx.from.username || null,
      firstName: draft.name,
      lastName: ctx.from.last_name || null
    },
    transactionLink: draft.transactionLink,
    transactionId: draft.transactionId,
    paymentMethod: draft.paymentMethod || null,
    paymentMethodLabel: draft.paymentMethodLabel || null
  };

  setRequest(store, requestId, request);
  clearDraft(store, ctx.from.id);
  await writeStore(store);

  await ctx.reply(buildPendingMessage(request));
  await notifyAdminAboutRequest(request);
}

async function handleDraftStep(ctx, text) {
  const store = await readStore();
  const draft = getDraft(store, ctx.from.id);

  if (!draft) {
    return false;
  }

  if (draft.step === 'package' || draft.step === 'paymentMethod') {
    return false;
  }

  if (draft.step === 'name') {
    draft.name = text;
    draft.step = 'transactionLink';
    setDraft(store, ctx.from.id, draft);
    await writeStore(store);
    await ctx.reply('Send the transaction link next.');
    return true;
  }

  if (draft.step === 'transactionLink') {
    const link = normalizeTransactionLink(text);
    if (draft.paymentMethod && !isValidTransactionLink(draft.paymentMethod, link)) {
      await ctx.reply(`That link does not look like a valid ${draft.paymentMethodLabel || 'receipt'} link. It should look like: ${getLinkFormatHint(draft.paymentMethod)}`);
      return true;
    }
    draft.transactionLink = link;
    draft.step = 'transactionId';
    setDraft(store, ctx.from.id, draft);
    await writeStore(store);
    await ctx.reply(`Send the transaction ID next (${getIdFormatHint(draft.paymentMethod)}).`);
    return true;
  }

  if (draft.step === 'transactionId') {
    const transactionId = normalizeTransactionId(text);
    if (draft.paymentMethod && !isValidTransactionId(draft.paymentMethod, transactionId)) {
      await ctx.reply(`That does not look like a valid ${draft.paymentMethodLabel || ''} transaction ID. It should look like: ${getIdFormatHint(draft.paymentMethod)}`);
      return true;
    }
    draft.transactionId = transactionId;
    await submitRequest(ctx, draft);
    return true;
  }

  return false;
}

bot.start(async (ctx) => {
  const packages = getPackageList();

  if (packages.length === 0) {
    await ctx.reply('No packages are configured yet. Add them in PACKAGES_JSON first.');
    return;
  }

  await ctx.reply(getStartMessage(packages), {
    reply_markup: buildPackageKeyboard(packages)
  });
});

bot.command('help', async (ctx) => {
  await ctx.reply([
    'Available commands:',
    '/buy - start the package and verification flow',
    '/myaccess - resend your approved voucher phrase',
    '/export - (admin) download the full Excel of approved transactions'
  ].join('\n'));
});

bot.command('buy', async (ctx) => {
  if (!ADMIN_CHAT_ID) {
    await ctx.reply('ADMIN_CHAT_ID is missing. Set it in .env before using /buy.');
    return;
  }

  const packages = getPackageList();
  if (packages.length === 0) {
    await ctx.reply('No packages are configured yet. Add them in PACKAGES_JSON first.');
    return;
  }

  const store = await readStore();
  setDraft(store, ctx.from.id, {
    step: 'package',
    startedAt: new Date().toISOString()
  });
  await writeStore(store);

  await ctx.reply(getPackageSelectionMessage(packages), {
    reply_markup: buildPackageKeyboard(packages)
  });
});

bot.command('cancel', async (ctx) => {
  const store = await readStore();
  clearDraft(store, ctx.from.id);
  await writeStore(store);
  await ctx.reply('Your current submission has been cancelled.');
});

bot.command('export', async (ctx) => {
  if (!ADMIN_CHAT_ID) {
    await ctx.reply('ADMIN_CHAT_ID is missing. Set it in .env before using /export.');
    return;
  }

  if (String(ctx.from.id) !== String(ADMIN_CHAT_ID)) {
    await ctx.reply('Only the admin can export the Excel file.');
    return;
  }

  try {
    await exportApprovedExcel(ctx.chat.id);
    await ctx.reply('Tap the button below to export the full Excel file anytime.', {
      reply_markup: {
        keyboard: [[{ text: 'Export Excel' }]],
        resize_keyboard: true
      }
    });
  } catch (error) {
    console.error('Failed to export workbook:', error);
    await ctx.reply('Export failed. Check the error log.');
  }
});

bot.action(/^(approve|reject):(.+)$/, async (ctx) => {
  const userIsAdmin = String(ctx.from.id) === String(ADMIN_CHAT_ID);
  if (!userIsAdmin) {
    await ctx.answerCbQuery('You are not allowed to review requests.');
    return;
  }

  try {
    const action = ctx.match[1];
    const requestId = ctx.match[2];
    const approved = action === 'approve';
    const result = await finalizeRequest(requestId, ctx.from.id, approved);

    if (!result) {
      await ctx.answerCbQuery('Request not found.');
      return;
    }

    await ctx.editMessageReplyMarkup({ inline_keyboard: [] });
    await ctx.answerCbQuery(approved ? 'Request approved.' : 'Request rejected.');
  } catch (error) {
    console.error('Failed to review request:', error);
    await ctx.answerCbQuery('Failed to process request.');
  }
});

bot.action(/^package:(.+)$/, async (ctx) => {
  const packageKey = ctx.match[1];
  const pkg = getPackageByKey(packageKey);

  if (!pkg) {
    await ctx.answerCbQuery('Package not found.');
    return;
  }

  const store = await readStore();
  setDraft(store, ctx.from.id, {
    step: 'paymentMethod',
    packageKey: pkg.key,
    packageLabel: pkg.label,
    priceCents: pkg.priceCents,
    currency: pkg.currency,
    startedAt: new Date().toISOString()
  });
  await writeStore(store);

  await ctx.answerCbQuery(`Selected ${pkg.label}`);
  await ctx.reply([
    `Package selected: ${pkg.label}`,
    `Price: ${formatMoney(pkg.priceCents, pkg.currency)}`,
    '',
    'Now choose your payment method:'
  ].join('\n'), {
    reply_markup: buildPaymentMethodKeyboard()
  });
});

bot.action(/^method:(.+)$/, async (ctx) => {
  const methodKey = ctx.match[1];
  const method = getPaymentMethodByKey(methodKey);

  if (!method) {
    await ctx.answerCbQuery('Unknown payment method.');
    return;
  }

  const store = await readStore();
  const draft = getDraft(store, ctx.from.id);

  if (!draft || draft.step !== 'paymentMethod') {
    await ctx.answerCbQuery('No active submission. Use /buy to start.');
    return;
  }

  draft.paymentMethod = method.key;
  draft.paymentMethodLabel = method.label;
  draft.step = 'name';
  setDraft(store, ctx.from.id, draft);
  await writeStore(store);

  await ctx.answerCbQuery(`Selected ${method.label}`);
  await ctx.reply('Send your full name to continue.');
});

bot.on('message', async (ctx) => {
  const text = ctx.message && ctx.message.text;
  if (!text) {
    return;
  }

  const trimmed = text.trim();

  if (trimmed === 'Export Excel') {
    if (String(ctx.from.id) === String(ADMIN_CHAT_ID)) {
      try {
        await exportApprovedExcel(ctx.chat.id);
      } catch (error) {
        console.error('Failed to export workbook:', error);
        await ctx.reply('Export failed. Check the error log.');
      }
    }
    return;
  }

  const handled = await handleDraftStep(ctx, trimmed);
  if (!handled && text.startsWith('/')) {
    return;
  }
});

bot.command('myaccess', async (ctx) => {
  const voucher = await getVoucherForUser(ctx.from.id);

  if (!voucher) {
    await ctx.reply('No approved voucher found for this account yet. Use /buy to submit your payment details.');
    return;
  }

  await ctx.reply(buildApprovedMessage({
    packageLabel: voucher.packageLabel || 'Your package',
    voucherPhrase: voucher.phrase
  }));
});

bot.catch((error) => {
  console.error('Bot error:', error);
});

(async () => {
  await ensureDataFile();
  if (!ADMIN_CHAT_ID) {
    console.warn('ADMIN_CHAT_ID is not set. The approval flow will not work until it is configured.');
  }
  await bot.launch();
  console.log('Tinat bot is running');
})();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
