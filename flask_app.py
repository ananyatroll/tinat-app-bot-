"""
Tinat Telegram bot - Python/Flask migration for PythonAnywhere Free Tier.

The bot runs entirely inside the request/response cycle of a Flask web worker.
Telegram is configured in webhook mode and POSTs every update to
https://<username>.pythonanywhere.com/telegram-webhook

Core features replicated from the Node.js bot:
  - /start package chooser (EUEE Preo, Freshman, UAT, University Department, Exit Exam)
  - state machine: awaiting package -> payment method (CBE/Telebirr) -> name -> transaction link -> transaction ID
  - admin approval with Approve / Reject inline buttons
  - voucher reservation from phrases/<pool>.txt, tracked in data/vouchers.json
  - Excel export of approved requests (one sheet per package) sent to the admin
  - redemption API so the Tinat app can redeem a voucher phrase once
"""

import base64
import copy
import json
import logging
import os
import secrets
from contextlib import contextmanager
from datetime import datetime, timezone

import requests
from flask import Flask, jsonify, request
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# Configuration / environment
# ---------------------------------------------------------------------------


def _load_dotenv(path=None):
    """Tiny .env loader (no extra dependency) so local testing works with
    the existing .env file. Real env vars always win."""
    path = path or os.path.join(BASE_DIR, '.env')
    try:
        with open(path, 'r', encoding='utf-8') as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, value = line.partition('=')
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except OSError:
        pass


_load_dotenv()

TELEGRAM_BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN') or os.environ.get('BOT_TOKEN', '')
ADMIN_CHAT_ID = os.environ.get('ADMIN_CHAT_ID', '')
PHRASES_DIR = os.environ.get('PHRASES_DIR', os.path.join(BASE_DIR, 'phrases'))
DATA_DIR = os.environ.get('DATA_DIR', os.path.join(BASE_DIR, 'data'))
EXPORTS_DIR = os.environ.get('EXPORTS_DIR', os.path.join(BASE_DIR, 'exports'))
USERS_FILE = os.environ.get('USERS_FILE', os.path.join(DATA_DIR, 'users.json'))
VOUCHERS_FILE = os.environ.get('VOUCHERS_FILE', os.path.join(DATA_DIR, 'vouchers.json'))
PROCESSED_UPDATES_FILE = os.environ.get('PROCESSED_UPDATES_FILE', os.path.join(DATA_DIR, 'processed_updates.json'))
MAX_RECENT_UPDATES = 500
UPDATE_DEDUPE_TTL_SECONDS = 60 * 60 * 24

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s: %(message)s',
)
logger = logging.getLogger('tinat-bot')

if not TELEGRAM_BOT_TOKEN:
    logger.warning('TELEGRAM_BOT_TOKEN is not set. The bot will not process messages.')
if not ADMIN_CHAT_ID:
    logger.warning('ADMIN_CHAT_ID is not set. The approval flow will not work.')


# ---------------------------------------------------------------------------
# Packages
# ---------------------------------------------------------------------------

DEFAULT_PACKAGES = [
    {'key': 'euee-preo', 'label': 'EUEE Preo', 'priceCents': 25000, 'currency': 'ETB', 'phrasePool': 'euee-preo'},
    {'key': 'freshman', 'label': 'Freshman', 'priceCents': 25000, 'currency': 'ETB', 'phrasePool': 'freshman'},
    {'key': 'uat', 'label': 'UAT', 'priceCents': 25000, 'currency': 'ETB', 'phrasePool': 'uat'},
    {'key': 'university-department', 'label': 'University Department', 'priceCents': 25000, 'currency': 'ETB', 'phrasePool': 'university-department'},
    {'key': 'exit-exam', 'label': 'Exit Exam', 'priceCents': 25000, 'currency': 'ETB', 'phrasePool': 'exit-exam'},
]


def load_packages():
    raw = os.environ.get('PACKAGES_JSON')
    if not raw:
        return copy.deepcopy(DEFAULT_PACKAGES)

    try:
        parsed = json.loads(raw)
    except ValueError:
        raise ValueError('PACKAGES_JSON must be valid JSON')

    if not isinstance(parsed, list) or not parsed:
        raise ValueError('PACKAGES_JSON must contain at least one package')

    packages = []
    for index, item in enumerate(parsed):
        key = str(item.get('key') or item.get('id') or 'package_%d' % (index + 1)).strip()
        label = str(item.get('label') or item.get('name') or 'Package %d' % (index + 1)).strip()
        try:
            price = int(float(item.get('priceCents', item.get('price_cents', 0))))
        except (TypeError, ValueError):
            price = 0
        currency = str(item.get('currency') or os.environ.get('CURRENCY', 'USD')).upper()
        pool = str(item.get('phrasePool') or item.get('pool') or item.get('key') or item.get('id') or key).strip()
        if key and label and price > 0:
            packages.append({
                'key': key,
                'label': label,
                'priceCents': price,
                'currency': currency,
                'phrasePool': pool,
            })
    return packages


PACKAGES = load_packages()


def get_package_by_key(package_key):
    for pkg in PACKAGES:
        if pkg['key'] == package_key:
            return pkg
    return None


def format_money(cents, currency='USD'):
    return '%.2f %s' % (cents / 100, currency)


def build_package_keyboard():
    rows = [[
        {'text': '%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])),
         'callback_data': 'package:%s' % pkg['key']}
    ] for pkg in PACKAGES]
    return {'inline_keyboard': rows}


def get_start_message():
    lines = [
        'Welcome to Tinat.',
        'Pick a package below, then I will ask for your name, transaction link, and transaction ID.',
        '',
    ]
    lines += ['%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])) for pkg in PACKAGES]
    lines += ['', 'After approval, the bot will send you one secret phrase for your selected package.']
    return '\n'.join(lines)


def get_package_selection_message():
    lines = [
        'Choose your package first:',
    ]
    lines += ['%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])) for pkg in PACKAGES]
    lines += ['', 'I will then ask for your name, transaction link, and transaction ID.']
    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# Payment methods
# ---------------------------------------------------------------------------

PAYMENT_METHODS = [
    {'key': 'cbe', 'label': 'Commercial Bank of Ethiopia'},
    {'key': 'telebirr', 'label': 'Telebirr'},
]

PAYMENT_LINK_PATTERNS = {
    'cbe': re.compile(r'^https://mbreciept\.cbe\.com\.et/\S+', re.IGNORECASE),
    'telebirr': re.compile(r'^https://transactioninfo\.ethiotelecom\.et/receipt/\S+', re.IGNORECASE),
}

PAYMENT_ID_PATTERNS = {
    'cbe': re.compile(r'^FT\d{4,}[A-Z0-9]{3,}$', re.IGNORECASE),
    'telebirr': re.compile(r'^[A-Z]{3}\d{1,}[A-Z0-9]{2,}$', re.IGNORECASE),
}


def is_valid_transaction_link(method_key, link):
    pattern = PAYMENT_LINK_PATTERNS.get(method_key)
    return bool(pattern and pattern.match(str(link or '').strip()))


def is_valid_transaction_id(method_key, value):
    pattern = PAYMENT_ID_PATTERNS.get(method_key)
    return bool(pattern and pattern.match(str(value or '').strip()))


def link_format_hint(method_key):
    if method_key == 'cbe':
        return 'https://mbreciept.cbe.com.et/<your-receipt-code>'
    if method_key == 'telebirr':
        return 'https://transactioninfo.ethiotelecom.et/receipt/<your-transaction-id>'
    return 'the full receipt link'


def id_format_hint(method_key):
    if method_key == 'cbe':
        return 'starts with FT, e.g. FT26222QKMBG'
    if method_key == 'telebirr':
        return 'letters and digits, e.g. DGJ22CMPJM'
    return 'your transaction ID'


def get_payment_method(method_key):
    for method in PAYMENT_METHODS:
        if method['key'] == method_key:
            return method
    return None


def build_payment_method_keyboard():
    rows = [[
        {'text': method['label'], 'callback_data': 'method:%s' % method['key']}
    ] for method in PAYMENT_METHODS]
    return {'inline_keyboard': rows}


def build_export_keyboard():
    return {
        'keyboard': [[{'text': 'Export Excel'}]],
        'resize_keyboard': True,
    }


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------


def utcnow():
    return datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')


def random_id(length):
    raw = base64.urlsafe_b64encode(secrets.token_bytes(length)).decode('ascii').rstrip('=')
    return raw[:length]


def _acquire_lock(fh):
    if os.name == 'nt':
        try:
            import msvcrt
            fh.seek(0)
            if not fh.read(1):
                fh.write(b'\0')
                fh.flush()
            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_LOCK, 1)
        except ImportError:
            pass
    else:
        try:
            import fcntl
            fcntl.flock(fh.fileno(), fcntl.LOCK_EX)
        except ImportError:
            pass


def _release_lock(fh):
    if os.name == 'nt':
        try:
            import msvcrt
            fh.seek(0)
            msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
        except Exception:
            pass
    else:
        try:
            import fcntl
            fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
        except Exception:
            pass


@contextmanager
def file_locked(path):
    """Cross-platform advisory file lock (fcntl on Linux/PythonAnywhere,
    msvcrt fallback on Windows for local testing)."""
    lock_path = path + '.lock'
    os.makedirs(os.path.dirname(lock_path) or '.', exist_ok=True)
    fh = open(lock_path, 'a+b')
    try:
        _acquire_lock(fh)
        yield
    finally:
        _release_lock(fh)
        fh.close()


def default_users():
    return {'users': {}, 'requests': {}, 'drafts': {}}


def default_vouchers():
    return {'packages': {}, 'issued': {}}


def read_json(path, default):
    with file_locked(path):
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except (OSError, ValueError):
            return copy.deepcopy(default)


def record_update(update_id):
    """Webhook delivery is at-least-once: Telegram retries a POST until it gets
    a 2xx, which would otherwise process the same update twice (advancing a
    draft two steps or double-submitting a request). Each update carries a
    monotonic update_id, so remember recent ones and skip repeats."""
    if update_id is None:
        return False

    def _record(store):
        key = str(update_id)
        if key in store:
            return True

        now = utcnow()
        store[key] = now

        cutoff = datetime.now(timezone.utc).timestamp() - UPDATE_DEDUPE_TTL_SECONDS
        for seen_key, seen_at in list(store.items()):
            try:
                ts = datetime.fromisoformat(str(seen_at).replace('Z', '+00:00')).timestamp()
            except ValueError:
                ts = 0
            if ts < cutoff:
                store.pop(seen_key, None)

        while len(store) > MAX_RECENT_UPDATES:
            store.pop(min(store, key=lambda k: store[k]), None)

        return False

    return mutate_json(PROCESSED_UPDATES_FILE, {}, _record)


def mutate_json(path, default, fn):
    """Atomically read -> mutate -> write a JSON file under a file lock.

    fn(data) may return anything; it is returned unchanged."""
    with file_locked(path):
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
        except (OSError, ValueError):
            data = copy.deepcopy(default)
        result = fn(data)
        os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
        tmp_path = path + '.tmp'
        with open(tmp_path, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp_path, path)
        return result


# ---------------------------------------------------------------------------
# Telegram API calls
# ---------------------------------------------------------------------------


def tg_call(method, timeout=30, **kwargs):
    if not TELEGRAM_BOT_TOKEN:
        return None
    url = 'https://api.telegram.org/bot%s/%s' % (TELEGRAM_BOT_TOKEN, method)
    try:
        response = requests.post(url, timeout=timeout, **kwargs)
        payload = response.json()
        if not payload.get('ok'):
            logger.warning('Telegram %s error: %s', method, payload)
        return payload
    except requests.RequestException as exc:
        logger.warning('Telegram %s network error: %s', method, exc)
        return None


def send_message(chat_id, text, reply_markup=None):
    params = {'chat_id': chat_id, 'text': text}
    if reply_markup is not None:
        params['reply_markup'] = reply_markup
    return tg_call('sendMessage', json=params)


def answer_callback_query(callback_query_id, text=None):
    params = {'callback_query_id': callback_query_id}
    if text:
        params['text'] = text
    return tg_call('answerCallbackQuery', json=params)


def edit_message_reply_markup(chat_id, message_id):
    params = {
        'chat_id': chat_id,
        'message_id': message_id,
        'reply_markup': {'inline_keyboard': []},
    }
    return tg_call('editMessageReplyMarkup', json=params)


def send_document(chat_id, file_path, caption=None):
    data = {'chat_id': chat_id}
    if caption:
        data['caption'] = caption
    with open(file_path, 'rb') as fh:
        files = {
            'document': (os.path.basename(file_path), fh,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }
        return tg_call('sendDocument', data=data, files=files)


# ---------------------------------------------------------------------------
# Message builders
# ---------------------------------------------------------------------------


def mask_transaction_id(value):
    value = str(value or '').strip()
    if len(value) <= 2:
        return '*' * len(value)
    return value[:2] + '*' * (len(value) - 2)


def build_admin_message(request):
    user = request.get('user') or {}
    full_name = ('%s %s' % (user.get('firstName') or '', user.get('lastName') or '')).strip()
    return '\n'.join([
        'New access request pending review:',
        'Request ID: %s' % request['requestId'],
        'Package: %s (%s)' % (request['packageLabel'], format_money(request['priceCents'], request['currency'])),
        'User: %s' % full_name,
        'Telegram: @%s (%s)' % (user.get('username') or 'no_username', user.get('id')),
        'Transaction ID: %s' % mask_transaction_id(request.get('transactionId')),
        'Transaction Link: %s' % request.get('transactionLink'),
        'Payment Method: %s' % (request.get('paymentMethodLabel') or request.get('paymentMethod') or 'N/A'),
        '',
        'Approve this request only after verifying the payment.',
    ])


def build_pending_message(request):
    return '\n'.join([
        'Your payment details were submitted.',
        'Request ID: %s' % request['requestId'],
        'I sent it to the admin for verification.',
        'You will get your voucher phrase after approval.',
    ])


def build_approved_message(package_label, voucher_phrase):
    return '\n'.join([
        'Payment approved. Your Tinat voucher phrase is below:',
        'Package: %s' % package_label,
        'Phrase: %s' % voucher_phrase,
        '',
        'Keep this phrase safe. If you lose it, send /myaccess in this bot.',
    ])


def notify_admin(request):
    if not ADMIN_CHAT_ID:
        raise RuntimeError('Missing ADMIN_CHAT_ID in environment')
    send_message(ADMIN_CHAT_ID, build_admin_message(request), reply_markup={
        'inline_keyboard': [[
            {'text': 'Approve', 'callback_data': 'approve:%s' % request['requestId']},
            {'text': 'Reject', 'callback_data': 'reject:%s' % request['requestId']},
        ]]
    })


# ---------------------------------------------------------------------------
# Voucher logic
# ---------------------------------------------------------------------------


class PoolEmptyError(Exception):
    """Raised when a phrase pool has no usable voucher left."""


def _load_pool_phrases(pool_key):
    path = os.path.join(PHRASES_DIR, pool_key + '.txt')
    try:
        with open(path, 'r', encoding='utf-8') as fh:
            return [line.strip() for line in fh if line.strip()]
    except OSError:
        return []


def _hydrate_pool(store, pkg):
    """Fill a package pool from phrases/<pool>.txt when it is empty,
    excluding phrases already handed out to anyone (global dedupe)."""
    pool_key = pkg['phrasePool']
    pool = store['packages'].get(pool_key)
    if pool is None:
        pool = {'available': [], 'issued': {}}
        store['packages'][pool_key] = pool
    pool.setdefault('available', [])
    pool.setdefault('issued', {})

    if pool['available']:
        return

    global_issued = {
        str(record.get('phrase') or '').strip()
        for record in (store.get('issued') or {}).values()
        if str(record.get('phrase') or '').strip()
    }
    pool['available'] = [p for p in _load_pool_phrases(pool_key) if p and p not in global_issued]


def reserve_voucher(request):
    """Reserve one unused phrase for an approved request. Runs under the
    voucher store file lock so two approvals can never take the same code."""
    pkg = get_package_by_key(request['packageKey'])
    if not pkg:
        raise ValueError('Unknown package: %s' % request['packageKey'])

    pool_key = pkg['phrasePool']
    request_id = request['requestId']
    user_id = str((request.get('user') or {}).get('id'))

    def _reserve(store):
        store.setdefault('packages', {})
        store.setdefault('issued', {})
        _hydrate_pool(store, pkg)

        pool = store['packages'][pool_key]
        available = pool.get('available') or []
        global_issued = {
            str(record.get('phrase') or '').strip()
            for record in (store.get('issued') or {}).values()
            if str(record.get('phrase') or '').strip()
        }

        phrase = ''
        while available:
            candidate = available.pop(0)
            if candidate and candidate not in global_issued:
                phrase = candidate
                break

        if not phrase:
            raise PoolEmptyError(
                'No voucher phrases left for package %s. Refill phrases/%s.txt'
                % (pkg['label'], pool_key)
            )

        record = {
            'requestId': request_id,
            'userId': user_id,
            'packageKey': pkg['key'],
            'packageLabel': pkg['label'],
            'packagePool': pool_key,
            'phrase': phrase,
            'status': 'reserved',
            'reservedAt': utcnow(),
        }
        pool['issued'][request_id] = record
        store['issued'][request_id] = record
        return record

    return mutate_json(VOUCHERS_FILE, default_vouchers, _reserve)


def get_voucher_for_user(user_id):
    store = read_json(VOUCHERS_FILE, default_vouchers())
    records = [
        record for record in (store.get('issued') or {}).values()
        if str(record.get('userId')) == str(user_id) and str(record.get('phrase') or '').strip()
    ]
    records.sort(key=lambda r: str(r.get('reservedAt') or ''), reverse=True)
    return records[0] if records else None


def redeem_phrase(phrase, user_id='', device_id=''):
    """Atomically mark a reserved phrase as redeemed. Shared by both the
    /api/redeem and /api/v1/vouchers/redeem endpoints."""
    phrase = str(phrase or '').strip()
    if not phrase:
        return {'ok': False, 'error': 'INVALID_PHRASE'}

    def _redeem(store):
        found_pool_key = None
        found_voucher = None
        for pool_key, pool in (store.get('packages') or {}).items():
            for voucher in (pool.get('issued') or {}).values():
                if str(voucher.get('phrase') or '').strip() == phrase:
                    found_pool_key = pool_key
                    found_voucher = voucher
                    break
            if found_voucher:
                break

        if not found_voucher:
            return {'ok': False, 'error': 'INVALID_PHRASE'}

        if found_voucher.get('status') == 'redeemed':
            return {'ok': False, 'error': 'ALREADY_REDEEMED'}

        now = utcnow()
        found_voucher['status'] = 'redeemed'
        found_voucher['redeemedAt'] = now
        found_voucher['redeemedByUserId'] = str(user_id or '')
        found_voucher['redeemedByDeviceId'] = str(device_id or '')
        store['issued'][found_voucher['requestId']] = found_voucher

        return {
            'ok': True,
            'packageKey': found_voucher.get('packageKey'),
            'packageLabel': found_voucher.get('packageLabel') or found_voucher.get('packageKey'),
            'redeemedAt': now,
            'access': {
                'packageKey': found_voucher.get('packageKey'),
                'enabled': True,
            },
        }

    return mutate_json(VOUCHERS_FILE, default_vouchers, _redeem)


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

EXPORT_HEADERS = [
    'Request ID', 'Package', 'Payment Method', 'Price', 'User Name',
    'Telegram Username', 'Telegram ID', 'Transaction ID', 'Transaction Link',
    'Voucher Phrase', 'Status', 'Approved At', 'Approved By',
]
EXPORT_WIDTHS = [20, 24, 18, 14, 22, 20, 16, 20, 40, 32, 14, 24, 16]


def build_approved_workbook(store):
    os.makedirs(EXPORTS_DIR, exist_ok=True)

    workbook = Workbook()
    workbook.remove(workbook.active)

    rows_by_sheet = {pkg['label']: [] for pkg in PACKAGES}
    rows_by_sheet['Other'] = []

    approved = [
        req for req in (store.get('requests') or {}).values()
        if req.get('status') == 'approved'
    ]
    approved.sort(key=lambda req: str(req.get('reviewedAt') or ''))

    for req in approved:
        user = req.get('user') or {}
        label = req.get('packageLabel')
        sheet_name = label if label in rows_by_sheet else 'Other'
        full_name = ('%s %s' % (user.get('firstName') or '', user.get('lastName') or '')).strip()
        rows_by_sheet[sheet_name].append([
            req.get('requestId'),
            req.get('packageLabel'),
            req.get('paymentMethodLabel') or req.get('paymentMethod') or '',
            format_money(req.get('priceCents', 0), req.get('currency')),
            full_name,
            user.get('username') or '',
            user.get('id'),
            req.get('transactionId'),
            req.get('transactionLink'),
            req.get('voucherPhrase') or '',
            req.get('status'),
            str(req.get('reviewedAt') or ''),
            str(req.get('reviewedBy') or ''),
        ])

    for sheet_name, rows in rows_by_sheet.items():
        if sheet_name == 'Other' and not rows:
            continue
        sheet = workbook.create_sheet(sheet_name)
        sheet.append(EXPORT_HEADERS)
        for index, width in enumerate(EXPORT_WIDTHS, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in rows:
            sheet.append(row)
        if rows:
            sheet.auto_filter.ref = sheet.dimensions

    filename = 'approved-%s.xlsx' % utcnow().replace(':', '-').replace('.', '-')
    file_path = os.path.join(EXPORTS_DIR, filename)
    workbook.save(file_path)
    return file_path


# ---------------------------------------------------------------------------
# State machine / request handling
# ---------------------------------------------------------------------------


def handle_draft_step(user_id, chat_id, text):
    def _step(store):
        draft = store['drafts'].get(user_id)
        if not draft:
            return ('none', None)

        step = draft.get('step')
        if not step or step in ('package', 'paymentMethod'):
            return ('none', None)

        if step == 'name':
            draft['name'] = text
            draft['step'] = 'transactionLink'
            return ('ask', 'Send the transaction link next.')

        if step == 'transactionLink':
            draft['transactionLink'] = text
            if not is_valid_transaction_link(draft.get('paymentMethod'), text):
                return ('ask', 'That link does not look like a valid %s link. It should look like: %s'
                        % (draft.get('paymentMethodLabel') or 'receipt', link_format_hint(draft.get('paymentMethod'))))
            draft['step'] = 'transactionId'
            return ('ask', 'Send the transaction ID next (%s).' % id_format_hint(draft.get('paymentMethod')))

        if step == 'transactionId':
            draft['transactionId'] = text
            if not is_valid_transaction_id(draft.get('paymentMethod'), text):
                return ('ask', 'That does not look like a valid %s transaction ID. It should look like: %s'
                        % (draft.get('paymentMethodLabel') or '', id_format_hint(draft.get('paymentMethod'))))
            return ('submit', draft)

        return ('none', None)

    code, payload = mutate_json(USERS_FILE, default_users, _step)

    if code == 'ask':
        send_message(chat_id, payload)
    elif code == 'submit':
        submit_request(user_id, chat_id, payload)


def submit_request(user_id, chat_id, draft):
    request_id = random_id(10)
    now = utcnow()

    def _submit(store):
        if user_id in store['users']:
            return ('already', None)

        request = {
            'requestId': request_id,
            'status': 'pending',
            'submittedAt': now,
            'packageKey': draft['packageKey'],
            'packageLabel': draft['packageLabel'],
            'priceCents': draft['priceCents'],
            'currency': draft['currency'],
            'user': {
                'id': user_id,
                'username': draft.get('username'),
                'firstName': draft.get('name'),
                'lastName': draft.get('last_name'),
            },
            'transactionLink': draft.get('transactionLink'),
            'transactionId': draft.get('transactionId'),
            'paymentMethod': draft.get('paymentMethod'),
            'paymentMethodLabel': draft.get('paymentMethodLabel'),
        }
        store['requests'][request_id] = request
        store['drafts'].pop(user_id, None)
        return ('submitted', request)

    code, request = mutate_json(USERS_FILE, default_users, _submit)

    if code == 'already':
        send_message(chat_id, 'You already have approved access. Use /myaccess to get your login details again.')
        return

    send_message(chat_id, build_pending_message(request))
    try:
        notify_admin(request)
    except Exception:
        logger.exception('Failed to notify admin for request %s', request_id)


def finalize_request(request_id, admin_user_id, approved):
    """Approve/reject a pending request. On approval, reserves the voucher
    atomically first so the request only flips to approved if a phrase is
    actually available."""
    if approved:
        try:
            def _approve(store):
                req = store['requests'].get(request_id)
                if not req or req.get('status') != 'pending':
                    return ('exists', req)
                voucher = reserve_voucher(req)
                now = utcnow()
                req['status'] = 'approved'
                req['reviewedAt'] = now
                req['reviewedBy'] = admin_user_id
                req['deliveredAt'] = now
                req['voucherPhrase'] = voucher['phrase']
                req['voucherStatus'] = 'delivered'
                req['excelExportAt'] = now
                user_id = str((req.get('user') or {}).get('id'))
                store['drafts'].pop(user_id, None)
                store['users'][user_id] = {'id': user_id, 'approvedAt': now}
                return ('approved', req)

            code, request = mutate_json(USERS_FILE, default_users, _approve)
        except PoolEmptyError as exc:
            logger.error('Pool empty for request %s: %s', request_id, exc)
            if ADMIN_CHAT_ID:
                send_message(ADMIN_CHAT_ID, 'Could not approve request %s: %s' % (request_id, exc))
            return ('pool_empty', None)
    else:
        def _reject(store):
            req = store['requests'].get(request_id)
            if not req or req.get('status') != 'pending':
                return ('exists', req)
            now = utcnow()
            req['status'] = 'rejected'
            req['reviewedAt'] = now
            req['reviewedBy'] = admin_user_id
            store['drafts'].pop(str((req.get('user') or {}).get('id')), None)
            return ('rejected', req)

        code, request = mutate_json(USERS_FILE, default_users, _reject)

    if code != 'approved' or not request:
        return code, request

    user_id = str((request.get('user') or {}).get('id'))

    try:
        send_message(user_id, build_approved_message(request['packageLabel'], request['voucherPhrase']))
    except Exception:
        logger.exception('Failed to send voucher to user %s', user_id)

    if ADMIN_CHAT_ID:
        try:
            store = read_json(USERS_FILE, default_users())
            count = sum(1 for r in (store.get('requests') or {}).values()
                        if r.get('status') == 'approved')
            workbook_path = build_approved_workbook(store)
            send_document(ADMIN_CHAT_ID, workbook_path,
                          caption='Excel updated: %d approved transaction(s) (incl. request %s). Use /export to re-download anytime.'
                                  % (count, request_id))
        except Exception:
            logger.exception('Failed to export/send workbook for request %s', request_id)

    return code, request


# ---------------------------------------------------------------------------
# Update handlers
# ---------------------------------------------------------------------------


def handle_message(msg):
    chat_id = (msg.get('chat') or {}).get('id')
    user = msg.get('from') or {}
    user_id = str(user.get('id'))
    text = str(msg.get('text') or '').strip()

    if not text or not chat_id or not user_id:
        return

    if text == '/start':
        if not PACKAGES:
            send_message(chat_id, 'No packages are configured yet. Add them in PACKAGES_JSON first.')
            return
        send_message(chat_id, get_start_message(), reply_markup=build_package_keyboard())
        return

    if text == '/help':
        send_message(chat_id,
                     'Available commands:\n'
                     '/buy - start the package and verification flow\n'
                     '/myaccess - resend your approved voucher phrase\n'
                     '/export - (admin) download the full Excel of approved transactions')
        return

    if text == '/buy':
        if not ADMIN_CHAT_ID:
            send_message(chat_id, 'ADMIN_CHAT_ID is missing. Set it before using /buy.')
            return
        if not PACKAGES:
            send_message(chat_id, 'No packages are configured yet. Add them in PACKAGES_JSON first.')
            return

        def _start(store):
            store['drafts'][user_id] = {'step': 'package', 'startedAt': utcnow()}

        mutate_json(USERS_FILE, default_users, _start)
        send_message(chat_id, get_package_selection_message(), reply_markup=build_package_keyboard())
        return

    if text == '/cancel':
        def _cancel(store):
            store['drafts'].pop(user_id, None)

        mutate_json(USERS_FILE, default_users, _cancel)
        send_message(chat_id, 'Your current submission has been cancelled.')
        return

    if text == '/myaccess':
        voucher = get_voucher_for_user(user_id)
        if not voucher:
            send_message(chat_id, 'No approved voucher found for this account yet. Use /buy to submit your payment details.')
            return
        send_message(chat_id, build_approved_message(
            voucher.get('packageLabel') or 'Your package', voucher['phrase']))
        return

    if text in ('/export', 'Export Excel'):
        if not ADMIN_CHAT_ID:
            send_message(chat_id, 'ADMIN_CHAT_ID is missing. Set it before using /export.')
            return
        if user_id != str(ADMIN_CHAT_ID):
            send_message(chat_id, 'Only the admin can export the Excel file.')
            return
        try:
            store = read_json(USERS_FILE, default_users())
            count = sum(1 for r in (store.get('requests') or {}).values()
                        if r.get('status') == 'approved')
            workbook_path = build_approved_workbook(store)
            send_document(chat_id, workbook_path,
                          caption='Full export of %d approved transaction(s).' % count)
            send_message(chat_id, 'Tap the button below to export the full Excel file anytime.',
                         reply_markup=build_export_keyboard())
        except Exception:
            logger.exception('Failed to export workbook on demand')
            send_message(chat_id, 'Export failed. Check the error log.')
        return

    if text.startswith('/'):
        return

    handle_draft_step(user_id, chat_id, text)


def handle_callback(cb):
    callback_id = cb['id']
    data = cb.get('data') or ''
    user = cb.get('from') or {}
    user_id = str(user.get('id'))
    message = cb.get('message') or {}
    chat_id = (message.get('chat') or {}).get('id')
    message_id = message.get('message_id')

    if data.startswith('package:'):
        package_key = data.split(':', 1)[1]
        pkg = get_package_by_key(package_key)
        if not pkg:
            answer_callback_query(callback_id, 'Package not found.')
            return

        def _select(store):
            store['drafts'][user_id] = {
                'step': 'paymentMethod',
                'packageKey': pkg['key'],
                'packageLabel': pkg['label'],
                'priceCents': pkg['priceCents'],
                'currency': pkg['currency'],
                'username': user.get('username'),
                'last_name': user.get('last_name'),
                'startedAt': utcnow(),
            }

        mutate_json(USERS_FILE, default_users, _select)
        answer_callback_query(callback_id, 'Selected %s' % pkg['label'])
        send_message(user_id,
                     'Package selected: %s\n'
                     'Price: %s\n'
                     '\n'
                     'Now choose your payment method:'
                     % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])),
                     reply_markup=build_payment_method_keyboard())
        return

    if data.startswith('method:'):
        method_key = data.split(':', 1)[1]
        method = get_payment_method(method_key)
        if not method:
            answer_callback_query(callback_id, 'Unknown payment method.')
            return

        def _set_method(store):
            draft = store['drafts'].get(user_id)
            if not draft or draft.get('step') != 'paymentMethod':
                return False
            draft['paymentMethod'] = method['key']
            draft['paymentMethodLabel'] = method['label']
            draft['step'] = 'name'
            return True

        updated = mutate_json(USERS_FILE, default_users, _set_method)
        if not updated:
            answer_callback_query(callback_id, 'No active submission. Use /buy to start.')
            return
        answer_callback_query(callback_id, 'Selected %s' % method['label'])
        send_message(user_id, 'Send your full name to continue.')
        return

    if data.startswith('approve:') or data.startswith('reject:'):
        action, request_id = data.split(':', 1)
        if user_id != str(ADMIN_CHAT_ID):
            answer_callback_query(callback_id, 'You are not allowed to review requests.')
            return

        try:
            code, _ = finalize_request(request_id, user_id, action == 'approve')
            if chat_id and message_id:
                edit_message_reply_markup(chat_id, message_id)
            if code == 'approved':
                answer_callback_query(callback_id, 'Request approved.')
            elif code == 'rejected':
                answer_callback_query(callback_id, 'Request rejected.')
            elif code == 'pool_empty':
                answer_callback_query(callback_id, 'Voucher pool empty for this package.')
            else:
                answer_callback_query(callback_id, 'Request not found or already processed.')
        except Exception:
            logger.exception('Failed to review request %s', request_id)
            answer_callback_query(callback_id, 'Failed to process request.')
        return

    answer_callback_query(callback_id, 'Unknown action.')


def handle_update(update):
    if 'callback_query' in update:
        handle_callback(update['callback_query'])
        return

    msg = update.get('message') or update.get('edited_message')
    if msg:
        handle_message(msg)


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

app = Flask(__name__)
application = app


@app.route('/')
def index():
    return 'Tinat bot webhook is live. POST Telegram updates to /telegram-webhook.'


@app.route('/health')
def health():
    return jsonify({'ok': True})


@app.route('/telegram-webhook', methods=['POST'])
def telegram_webhook():
    if not request.is_json:
        return jsonify({'ok': False, 'error': 'expected JSON'}), 400

    update = request.get_json(silent=True) or {}
    if record_update(update.get('update_id')):
        return jsonify({'ok': True})
    try:
        handle_update(update)
    except Exception:
        logger.exception('Failed to process update')
    # Always 200 so Telegram stops retrying; failures are visible in the
    # PythonAnywhere error log and get retried manually if needed.
    return jsonify({'ok': True})


@app.route('/api/redeem', methods=['POST'])
def api_redeem():
    body = request.get_json(silent=True) or {}
    result = redeem_phrase(body.get('phrase'), body.get('userId'), body.get('deviceId'))
    if result.get('ok'):
        return jsonify({
            'status': 'success',
            'packageKey': result['packageKey'],
            'packageLabel': result['packageLabel'],
        }), 200
    return jsonify({'status': 'invalid', 'error': result.get('error')}), 400


@app.route('/api/v1/vouchers/redeem', methods=['POST'])
def api_redeem_v1():
    """Legacy contract from docs/voucher-redemption-contract.md."""
    body = request.get_json(silent=True) or {}
    result = redeem_phrase(body.get('phrase'), body.get('userId'), body.get('deviceId'))
    if result.get('ok'):
        return jsonify(result), 200
    return jsonify(result), 400


def register_webhook(webhook_url):
    """Set the Telegram webhook. Call from a PythonAnywhere Bash console:
    python -c "from flask_app import register_webhook; register_webhook('https://<user>.pythonanywhere.com/telegram-webhook')"
    """
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError('TELEGRAM_BOT_TOKEN is not set')
    response = requests.post(
        'https://api.telegram.org/bot%s/setWebhook' % TELEGRAM_BOT_TOKEN,
        json={'url': webhook_url, 'allowed_updates': ['message', 'callback_query', 'edited_message']},
        timeout=30,
    )
    payload = response.json()
    logger.info('setWebhook -> %s', payload)
    if not payload.get('ok'):
        raise RuntimeError('setWebhook failed: %s' % payload)
    return payload


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
