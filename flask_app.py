"""
Tinat Telegram bot - Python/Flask migration for PythonAnywhere Free Tier.

The bot runs entirely inside the request/response cycle of a Flask web worker.
Telegram is configured in webhook mode and POSTs every update to
https://<username>.pythonanywhere.com/telegram-webhook

Feature parity with the current Node.js bot:
   - /start package chooser (EUEE Prep, Freshman, UAT, University Department, Exit Exam)
  - state machine: package -> phone (verified via Telegram contact share) ->
    payment method (CBE/Telebirr) -> name -> transaction link -> transaction ID
  - admin approval with Approve / Reject inline buttons
  - Approve immediately reserves one unused voucher from the correct package
    pool and sends the phrase to the buyer; Reject marks the payment rejected
  - support role (/pending or the Assign buttons) for assigning vouchers to
    already-approved purchases that are missing one
  - voucher assignment tracked in data/vouchers.json with owner / phone /
    transaction / assigned-by audit fields
  - Excel export of approved requests (one sheet per package) sent to the admin
  - redemption API that verifies Telegram Mini App initData, voucher ownership,
    payment approval and voucher state, then atomically marks the voucher
    REDEEMED and writes the entitlement to data/entitlements.json
"""

import base64
import copy
import hashlib
import hmac
import json
import logging
import os
import re
import secrets
import time
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from urllib.parse import unquote

import requests
from flask import Flask, jsonify, request
from openpyxl import Workbook
from openpyxl.styles import Font

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


def _abs_path(value, default):
    """Resolve a possibly-relative configured path against BASE_DIR so the
    app works regardless of the process working directory (important on
    PythonAnywhere, where the WSGI working directory is configurable)."""
    value = (value or '').strip() or default
    if not os.path.isabs(value):
        value = os.path.join(BASE_DIR, value)
    return os.path.normpath(value)


_load_dotenv()

TELEGRAM_BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN') or os.environ.get('BOT_TOKEN', '')
ADMIN_CHAT_ID = os.environ.get('ADMIN_CHAT_ID', '')
SUPPORT_CHAT_ID = os.environ.get('SUPPORT_CHAT_ID') or ADMIN_CHAT_ID or ''
DATA_DIR = _abs_path(os.environ.get('DATA_DIR'), os.path.join(BASE_DIR, 'data'))
PHRASES_DIR = _abs_path(os.environ.get('PHRASES_DIR'), os.path.join(BASE_DIR, 'phrases'))
EXPORTS_DIR = _abs_path(os.environ.get('EXPORTS_DIR'), os.path.join(BASE_DIR, 'exports'))
USERS_FILE = _abs_path(os.environ.get('USERS_FILE'), os.path.join(DATA_DIR, 'users.json'))
VOUCHERS_FILE = _abs_path(os.environ.get('VOUCHERS_FILE'), os.path.join(DATA_DIR, 'vouchers.json'))
ENTITLEMENTS_FILE = _abs_path(os.environ.get('ENTITLEMENTS_FILE'), os.path.join(DATA_DIR, 'entitlements.json'))
PROCESSED_UPDATES_FILE = _abs_path(os.environ.get('PROCESSED_UPDATES_FILE'), os.path.join(DATA_DIR, 'processed_updates.json'))
MAX_RECENT_UPDATES = 500
UPDATE_DEDUPE_TTL_SECONDS = 60 * 60 * 24

REDEEM_RATE_LIMIT_MAX = int(os.environ.get('REDEEM_RATE_LIMIT_MAX', '5'))
REDEEM_RATE_LIMIT_WINDOW_MS = int(os.environ.get('REDEEM_RATE_LIMIT_WINDOW_MS', str(10 * 60 * 1000)))
INIT_DATA_MAX_AGE_SECONDS = int(os.environ.get('INIT_DATA_MAX_AGE_SECONDS', str(24 * 60 * 60)))
AUTH_TOKEN_TTL_SECONDS = int(os.environ.get('AUTH_TOKEN_TTL_SECONDS', str(24 * 60 * 60)))
AUTH_TOKENS_FILE = _abs_path(os.environ.get('AUTH_TOKENS_FILE'), os.path.join(DATA_DIR, 'auth_tokens.json'))

ACCESS_TOKENS_FILE = _abs_path(os.environ.get('ACCESS_TOKENS_FILE'), os.path.join(DATA_DIR, 'access_tokens.json'))
ACCESS_TOKEN_TTL_SECONDS = int(os.environ.get('ACCESS_TOKEN_TTL_SECONDS', str(365 * 24 * 60 * 60)))
ANDROID_RATE_LIMIT_MAX = int(os.environ.get('ANDROID_RATE_LIMIT_MAX', str(REDEEM_RATE_LIMIT_MAX)))
ANDROID_RATE_LIMIT_WINDOW_MS = int(os.environ.get('ANDROID_RATE_LIMIT_WINDOW_MS', str(REDEEM_RATE_LIMIT_WINDOW_MS)))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s: %(message)s',
)
logger = logging.getLogger('tinat-bot')

if not TELEGRAM_BOT_TOKEN:
    logger.warning('TELEGRAM_BOT_TOKEN is not set. The bot will not process messages.')
if not ADMIN_CHAT_ID:
    logger.warning('ADMIN_CHAT_ID is not set. The approval flow will not work.')


def _prepare_data_dir():
    """Validate the persistent data directory at startup.

    Logs the resolved DATA_DIR and its writability, and refuses to start
    when an explicitly-configured directory outside the project (for example
    Railway's volume at /data) is missing or not writable. It never
    auto-creates a mount point: doing so would silently persist to ephemeral
    storage and lose data on the next redeploy. The project-local default
    (BASE_DIR/data) is created lazily for development and tests.

    This is intentionally strict when DATA_DIR comes from the environment:
    a missing /data means the volume is not actually mounted, and the app
    must surface that instead of quietly running on throwaway storage.
    """
    inside_project = os.path.normpath(DATA_DIR).startswith(
        os.path.normpath(BASE_DIR) + os.sep)

    if os.path.isdir(DATA_DIR):
        probe = os.path.join(DATA_DIR, '.write-test')
        try:
            with open(probe, 'w', encoding='utf-8') as fh:
                fh.write('ok')
            try:
                os.remove(probe)
            except OSError:
                pass
        except OSError as exc:
            raise RuntimeError(
                'DATA_DIR %r is configured but not writable: %s' % (DATA_DIR, exc))
        logger.info('DATA_DIR: %s (writable)', DATA_DIR)
        return

    if not inside_project:
        raise RuntimeError(
            'DATA_DIR %r is configured but the directory does not exist. '
            'On Railway this means the persistent volume is not mounted at '
            'that path. The app will NOT fall back to ephemeral storage; '
            'mount the volume (or fix DATA_DIR) and restart.' % DATA_DIR)

    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except OSError as exc:
        raise RuntimeError('Cannot create DATA_DIR %r: %s' % (DATA_DIR, exc))
    logger.info('DATA_DIR: %s (created)', DATA_DIR)


_prepare_data_dir()

for _storage_label, _storage_path in (
        ('users', USERS_FILE),
        ('vouchers', VOUCHERS_FILE),
        ('entitlements', ENTITLEMENTS_FILE),
        ('processed_updates', PROCESSED_UPDATES_FILE),
        ('auth_tokens', AUTH_TOKENS_FILE),
        ('access_tokens', ACCESS_TOKENS_FILE)):
    logger.info('Storage file %s: %s', _storage_label, _storage_path)


# ---------------------------------------------------------------------------
# Packages
# ---------------------------------------------------------------------------

DEFAULT_PACKAGES = [
    {'key': 'euee-prep', 'label': 'EUEE Prep', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'euee-prep'},
    {'key': 'freshman', 'label': 'Freshman', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'freshman'},
    {'key': 'uat', 'label': 'UAT', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'uat'},
    {'key': 'university-department', 'label': 'University Department', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'university-department'},
    {'key': 'coc', 'label': '📋 COC Exam Preparation', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'exit-exam'},
    {'key': 'exit-exam', 'label': 'Exit Exam', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'exit-exam'},
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
        'Pick a package below, then I will ask you to share your phone number, payment method, name, transaction link, and transaction ID.',
        '',
    ]
    lines += ['%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])) for pkg in PACKAGES]
    lines += ['', 'After approval, our support team sends you one secret phrase for your selected package.']
    return '\n'.join(lines)


def get_package_selection_message():
    lines = [
        'Choose your package first:',
    ]
    lines += ['%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])) for pkg in PACKAGES]
    lines += ['', 'I will then ask you to share your phone number, payment method, name, transaction link, and transaction ID.']
    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# Payment methods
# ---------------------------------------------------------------------------

PAYMENT_METHODS = [
    {'key': 'cbe', 'label': 'Commercial Bank of Ethiopia (CBE)'},
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
        return 'letters and digits, e.g. DG*********, DH********'
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


def build_phone_share_keyboard():
    return {
        'keyboard': [[{'text': 'Share Phone Number', 'request_contact': True}]],
        'resize_keyboard': True,
        'one_time_keyboard': True,
    }


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------


def utcnow(offset_seconds=0):
    return (datetime.now(timezone.utc) + timedelta(seconds=offset_seconds)).isoformat().replace('+00:00', 'Z')


def random_id(length):
    raw = base64.urlsafe_b64encode(secrets.token_bytes(length)).decode('ascii').rstrip('=')
    return raw[:length]


def is_plausible_phone(value):
    digits = re.sub(r'\D', '', str(value or ''))
    return 9 <= len(digits) <= 15


def normalize_phone(value):
    """Canonical phone form: digits only (no +, spaces, dashes)."""
    return re.sub(r'\D', '', str(value or ''))


def mask_phone(value):
    """Mask a phone for logs, keeping only the last 4 digits."""
    digits = normalize_phone(value)
    if len(digits) <= 4:
        return '***'
    return '*' * (len(digits) - 4) + digits[-4:]


def _acquire_lock(fh, timeout=30):
    if os.name == 'nt':
        try:
            import msvcrt
            deadline = time.time() + timeout
            while True:
                try:
                    fh.seek(0)
                    if not fh.read(1):
                        fh.write(b'\0')
                        fh.flush()
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_LOCK, 1)
                    break
                except OSError:
                    # msvcrt.locking is non-blocking on Windows: both locking
                    # AND reading the locked byte from another handle raise
                    # PermissionError while another thread holds the lock.
                    # Poll until the region is free so concurrent writers
                    # serialize the same as flock on Linux.
                    if time.time() > deadline:
                        raise
                    time.sleep(0.05)
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


def default_entitlements():
    return {'entitlements': {}}


def read_json(path, default):
    with file_locked(path):
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except (OSError, ValueError):
            return copy.deepcopy(default)


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
        # os.replace can transiently fail on Windows when the destination is
        # still open (AV or OneDrive sync); retry briefly instead of crashing.
        deadline = time.time() + 10
        while True:
            try:
                os.replace(tmp_path, path)
                break
            except OSError:
                if time.time() > deadline:
                    raise
                time.sleep(0.05)
        return result


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


def send_message(chat_id, text, reply_markup=None, parse_mode=None):
    params = {'chat_id': chat_id, 'text': text}
    if reply_markup is not None:
        params['reply_markup'] = reply_markup
    if parse_mode:
        params['parse_mode'] = parse_mode
    return tg_call('sendMessage', json=params)


def answer_callback_query(callback_query_id, text=None):
    params = {'callback_query_id': callback_query_id}
    if text:
        params['text'] = text
    return tg_call('answerCallbackQuery', json=params)


def edit_message_reply_markup(chat_id, message_id, reply_markup=None):
    params = {
        'chat_id': chat_id,
        'message_id': message_id,
        'reply_markup': reply_markup or {'inline_keyboard': []},
    }
    return tg_call('editMessageReplyMarkup', json=params)


def edit_message_text(chat_id, message_id, text, reply_markup=None, parse_mode=None):
    params = {
        'chat_id': chat_id,
        'message_id': message_id,
        'text': text,
    }
    if reply_markup is not None:
        params['reply_markup'] = reply_markup
    if parse_mode:
        params['parse_mode'] = parse_mode
    return tg_call('editMessageText', json=params)


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
# Roles
# ---------------------------------------------------------------------------


def is_admin(user_id):
    return str(user_id) == str(ADMIN_CHAT_ID)


def is_support(user_id):
    return bool(SUPPORT_CHAT_ID) and str(user_id) == str(SUPPORT_CHAT_ID)


def can_assign_vouchers(user_id):
    return is_admin(user_id) or is_support(user_id)


# ---------------------------------------------------------------------------
# Canonical Product Catalog & Entitlement Mappings
# Architecture: Freshman, University, and COC Exam
# ---------------------------------------------------------------------------

CANONICAL_PRODUCTS = {
    # Freshman Natural Science
    'freshman_natural_science_y1_sem1': {
        'id': 'freshman_natural_science_y1_sem1',
        'category': 'freshman',
        'label': 'Freshman Natural Science - Year 1 Semester 1',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['freshman_natural_science_y1_sem1'],
        'phrasePool': 'freshman',
    },
    'freshman_natural_science_y1_sem2': {
        'id': 'freshman_natural_science_y1_sem2',
        'category': 'freshman',
        'label': 'Freshman Natural Science - Year 1 Semester 2',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['freshman_natural_science_y1_sem2'],
        'phrasePool': 'freshman',
    },
    'freshman_natural_science_y1_full_year': {
        'id': 'freshman_natural_science_y1_full_year',
        'category': 'freshman',
        'label': 'Freshman Natural Science - Year 1 Full Academic Year',
        'priceCents': 50000,
        'currency': 'ETB',
        'entitlements': ['freshman_natural_science_y1_sem1', 'freshman_natural_science_y1_sem2'],
        'phrasePool': 'freshman',
    },

    # Freshman Social Science
    'freshman_social_science_y1_sem1': {
        'id': 'freshman_social_science_y1_sem1',
        'category': 'freshman',
        'label': 'Freshman Social Science - Year 1 Semester 1',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['freshman_social_science_y1_sem1'],
        'phrasePool': 'freshman',
    },
    'freshman_social_science_y1_sem2': {
        'id': 'freshman_social_science_y1_sem2',
        'category': 'freshman',
        'label': 'Freshman Social Science - Year 1 Semester 2',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['freshman_social_science_y1_sem2'],
        'phrasePool': 'freshman',
    },
    'freshman_social_science_y1_full_year': {
        'id': 'freshman_social_science_y1_full_year',
        'category': 'freshman',
        'label': 'Freshman Social Science - Year 1 Full Academic Year',
        'priceCents': 50000,
        'currency': 'ETB',
        'entitlements': ['freshman_social_science_y1_sem1', 'freshman_social_science_y1_sem2'],
        'phrasePool': 'freshman',
    },

    # University Computer Science (Year 2 example)
    'university_computer_science_y2_sem1': {
        'id': 'university_computer_science_y2_sem1',
        'category': 'university',
        'label': 'University Computer Science - Year 2 Semester 1',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['university_computer_science_y2_sem1'],
        'phrasePool': 'university-department',
    },
    'university_computer_science_y2_sem2': {
        'id': 'university_computer_science_y2_sem2',
        'category': 'university',
        'label': 'University Computer Science - Year 2 Semester 2',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['university_computer_science_y2_sem2'],
        'phrasePool': 'university-department',
    },
    'university_computer_science_y2_full_year': {
        'id': 'university_computer_science_y2_full_year',
        'category': 'university',
        'label': 'University Computer Science - Year 2 Full Academic Year',
        'priceCents': 50000,
        'currency': 'ETB',
        'entitlements': ['university_computer_science_y2_sem1', 'university_computer_science_y2_sem2'],
        'phrasePool': 'university-department',
    },

    # Generic University Department
    'university-department': {
        'id': 'university-department',
        'category': 'university',
        'label': 'University Department',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['university-department'],
        'phrasePool': 'university-department',
    },
    'euee-prep': {
        'id': 'euee-prep',
        'category': 'freshman',
        'label': 'EUEE Prep',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['euee-prep'],
        'phrasePool': 'euee-prep',
    },
    'freshman': {
        'id': 'freshman',
        'category': 'freshman',
        'label': 'Freshman',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['freshman'],
        'phrasePool': 'freshman',
    },
    'uat': {
        'id': 'uat',
        'category': 'freshman',
        'label': 'UAT',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['uat'],
        'phrasePool': 'uat',
    },
    'exit-exam': {
        'id': 'exit-exam',
        'category': 'university',
        'label': 'Exit Exam',
        'priceCents': 30000,
        'currency': 'ETB',
        'entitlements': ['exit-exam'],
        'phrasePool': 'exit-exam',
    },

    # COC Exam Category
    'coc_medical': {
        'id': 'coc_medical',
        'category': 'coc',
        'label': 'Medical COC Exam Preparation',
        'priceCents': 40000,
        'currency': 'ETB',
        'entitlements': ['coc_medical'],
        'phrasePool': 'exit-exam',
    },
    'coc_law': {
        'id': 'coc_law',
        'category': 'coc',
        'label': 'Law COC Exam Preparation',
        'priceCents': 40000,
        'currency': 'ETB',
        'entitlements': ['coc_law'],
        'phrasePool': 'exit-exam',
    },
    'coc_engineering': {
        'id': 'coc_engineering',
        'category': 'coc',
        'label': 'Engineering COC Exam Preparation',
        'priceCents': 40000,
        'currency': 'ETB',
        'entitlements': ['coc_engineering'],
        'phrasePool': 'exit-exam',
    },
}

PURCHASES_FILE = _abs_path(os.environ.get('PURCHASES_FILE'), os.path.join(DATA_DIR, 'purchases.json'))
PACKAGES_FILE = _abs_path(os.environ.get('PACKAGES_FILE'), os.path.join(DATA_DIR, 'packages.json'))

def default_purchases_store():
    return {'purchases': {}}

def get_product_config(product_id):
    if not product_id:
        return None
    product_id = str(product_id).strip()
    if product_id in CANONICAL_PRODUCTS:
        return copy.deepcopy(CANONICAL_PRODUCTS[product_id])
    
    # Generic format handling for university_{dept}_y{year}_{sem/full}
    if product_id.startswith('university_'):
        is_full = product_id.endswith('_full_year')
        sem1 = product_id.replace('_full_year', '_sem1')
        sem2 = product_id.replace('_full_year', '_sem2')
        return {
            'id': product_id,
            'category': 'university',
            'label': product_id.replace('_', ' ').title(),
            'priceCents': 50000 if is_full else 30000,
            'currency': 'ETB',
            'entitlements': [sem1, sem2] if is_full else [product_id],
            'phrasePool': 'university-department',
        }
    
    # Fallback to load_packages
    for pkg in load_packages():
        if pkg['key'] == product_id:
            return {
                'id': pkg['key'],
                'category': 'general',
                'label': pkg['label'],
                'priceCents': pkg['priceCents'],
                'currency': pkg.get('currency', 'ETB'),
                'phrasePool': pkg.get('phrasePool', pkg['key']),
            }
    return None

def get_product_price(product_id):
    cfg = get_product_config(product_id)
    return cfg['priceCents'] if cfg else 30000

def get_product_entitlements(product_id):
    cfg = get_product_config(product_id)
    return cfg['entitlements'] if cfg else [product_id]

# ---------------------------------------------------------------------------
# Purchase References Storage & Management
# ---------------------------------------------------------------------------

def generate_purchase_reference():
    chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ'
    code = ''.join(secrets.choice(chars) for _ in range(5))
    return 'TH-%s' % code

def create_purchase_request(product_id, user_id=None, user_phone=None):
    cfg = get_product_config(product_id)
    if not cfg:
        raise ValueError('Invalid product_id: %s' % product_id)
    
    ref = generate_purchase_reference()
    now = utcnow()
    purchase = {
        'purchaseReference': ref,
        'productId': cfg['id'],
        'productLabel': cfg['label'],
        'category': cfg['category'],
        'priceCents': cfg['priceCents'],
        'currency': cfg['currency'],
        'entitlements': cfg['entitlements'],
        'phrasePool': cfg['phrasePool'],
        'userId': str(user_id) if user_id else None,
        'userPhone': normalize_phone(user_phone) if user_phone else None,
        'status': 'pending',
        'createdAt': now,
        'updatedAt': now,
    }
    def _mutate(data):
        purchases = data.setdefault('purchases', {})
        purchases[ref] = purchase
        return purchase
    mutate_json(PURCHASES_FILE, default_purchases_store(), _mutate)
    return ref

def get_purchase_by_reference(ref):
    if not ref:
        return None
    ref = str(ref).strip().upper()
    if not ref.startswith('TH-'):
        ref = 'TH-' + ref.lstrip('-')
    store = read_json(PURCHASES_FILE, default_purchases_store())
    return (store.get('purchases') or {}).get(ref)

def is_valid_purchase_reference_format(ref):
    if not ref:
        return False
    ref = str(ref).strip().upper()
    return bool(re.match(r'^(TH-)?[A-Z0-9]{4,8}$', ref))

# ---------------------------------------------------------------------------
# Packages & Dynamic Management
# ---------------------------------------------------------------------------

DEFAULT_PACKAGES = [
    {'key': 'euee-prep', 'label': 'EUEE Prep', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'euee-prep'},
    {'key': 'freshman', 'label': 'Freshman', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'freshman'},
    {'key': 'uat', 'label': 'UAT', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'uat'},
    {'key': 'university-department', 'label': 'University Department', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'university-department'},
    {'key': 'exit-exam', 'label': 'Exit Exam', 'priceCents': 30000, 'currency': 'ETB', 'phrasePool': 'exit-exam'},
]

def default_packages_store():
    return {'packages': copy.deepcopy(DEFAULT_PACKAGES)}

def load_packages():
    raw_env = os.environ.get('PACKAGES_JSON')
    if raw_env:
        try:
            parsed = json.loads(raw_env)
            if isinstance(parsed, list) and parsed:
                return parsed
        except ValueError:
            pass

    store = read_json(PACKAGES_FILE, default_packages_store())
    packages = store.get('packages')
    if not packages or not isinstance(packages, list):
        return copy.deepcopy(DEFAULT_PACKAGES)
    return packages

def save_packages(packages_list):
    def _mutate(data):
        data['packages'] = packages_list
        return True
    mutate_json(PACKAGES_FILE, default_packages_store(), _mutate)

def get_package_by_key(package_key):
    cfg = get_product_config(package_key)
    if cfg:
        return {
            'key': cfg['id'],
            'label': cfg['label'],
            'priceCents': cfg['priceCents'],
            'currency': cfg['currency'],
            'phrasePool': cfg['phrasePool'],
        }
    for pkg in load_packages():
        if pkg['key'] == package_key:
            return pkg
    return None

def format_money(cents, currency='ETB'):
    return '%.2f %s' % (cents / 100, currency)

# ---------------------------------------------------------------------------
# Multilingual Support (5 Languages)
# Languages: English, Amharic, Oromiffa, Somali, Tigrinya
# ---------------------------------------------------------------------------

LANGUAGES = {
    'en': 'English 🇬🇧',
    'am': 'አማርኛ 🇪🇹',
    'om': 'Afaan Oromoo 🇪🇹',
    'so': 'Af-Soomaali 🇸🇴',
    'ti': 'ትግርኛ 🇪🇹',
}

MESSAGES = {
    'en': {
        'lang_select': 'Please select your preferred language:',
        'start': 'Welcome to Temhiro Study App! Please choose a package to continue:',
        'package_chosen': 'Package: %s\nPrice: %s\nPlease share your phone number using the button below:',
        'share_phone_btn': '📱 Share Phone Number',
        'invalid_phone': 'Please use the official "Share Phone Number" button below.',
        'payment_method': 'How do you want to pay? Please select your payment method:',
        'invalid_method': 'Invalid payment method. Please select using the buttons below.',
        'ask_name': 'Great. What is your full name?',
        'ask_link': 'Send the full transaction receipt link of your payment.\nFormat: %s',
        'invalid_link': 'That link does not look like a valid %s receipt.\nExpected format: %s\nPlease send the full link again.',
        'ask_txid': 'Now send your transaction ID / Ref Number.\nFormat: %s',
        'invalid_txid': 'That transaction ID does not look valid for %s.\nExpected format: %s\nPlease try again.',
        'submitted_pending': 'Thank you! Your payment details were submitted for review (Request ID: %s). You will receive your voucher code here after approval.',
        'already_submitted': 'Your purchase is currently pending admin review. Send /myaccess to check your status.',
        'already_approved': 'You already have an approved purchase (Request ID: %s). Send /myaccess to view your voucher code.',
        'myaccess_redeemed': 'Your voucher was redeemed on %s.',
        'myaccess_voucher': 'Your voucher for %s:\nPhrase: %s',
        'no_voucher': 'No active voucher found. Use /start to select a package and subscribe.',
        'cancelled': 'Operation cancelled. Send /start to begin again.',
        'approved_msg': '🎉 *Your payment for %s has been approved!*\n\n📱 *Phone Number:*\n`%s` \n\n🔑 *Redeem Code:*\n`%s`\n\nTo redeem, open the Temhiro Study App and enter your phone number and redeem code.',
        'copy_phone': '📱 %s',
        'copy_code': '🔑 %s',
        'btn_copy_phone': '📋 Copy Phone Number',
        'btn_copy_code': '🔑 Copy Redeem Code',
        'rejected_msg': '❌ Your payment request for %s was rejected by the admin (Request ID: %s). If you believe this is an error, please contact support.',
        'approved_pending_pool': 'Your payment for %s was approved! Our support team is assigning your voucher phrase shortly (Request ID: %s).',
    },
    'am': {
        'lang_select': 'እባክዎ የሚመርጡትን ቋንቋ ይምረጡ:',
        'start': 'እንኳን ወደ ተምህሮ የጥናት መተግበሪያ በደህና መጡ! ለመቀጠል እባክዎ ፓኬጅ ይምረጡ:',
        'package_chosen': 'ፓኬጅ: %s\nዋጋ: %s\nእባክዎ ከታች ያለውን ቁልፍ ተጠቅመው ስልክ ቁጥርዎን ያጋሩ:',
        'share_phone_btn': '📱 ስልክ ቁጥር ያጋሩ',
        'invalid_phone': 'እባክዎ ከታች ያለውን ይፋዊ "ስልክ ቁጥር ያጋሩ" ቁልፍ ይጠቀሙ።',
        'payment_method': 'እንዴት መክፈል ይፈልጋሉ? እባክዎ የክፍያ መንገድ ይምረጡ:',
        'invalid_method': 'የተሳሳተ የክፍያ መንገድ። እባክዎ ከታች ያሉትን ቁልፎች ይጠቀሙ።',
        'ask_name': 'በጣም ጥሩ። ሙሉ ስምዎን ያስገቡ:',
        'ask_link': 'የክፍያዎን ደረሰኝ ሙሉ ሊንክ ይላኩ።\nቅርጸት: %s',
        'invalid_link': 'የላኩት ሊንክ የ %s ደረሰኝ አይመስልም።\nየሚጠበቀው ቅርጸት: %s\nእባክዎ ሙሉውን ሊንክ እንደገና ይላኩ።',
        'ask_txid': 'አሁን የግብይት መለያ ቁጥርዎን (Transaction ID / Ref Number) ይላኩ።\nቅርጸት: %s',
        'invalid_txid': 'ያስገቡት የግብይት መለያ ቁጥር ለ %s ትክክለኛ አይመስልም።\nየሚጠበቀው ቅርጸት: %s\nእባክዎ እንደገና ይሞክሩ።',
        'submitted_pending': 'እናመሰግናለን! የክፍያ ዝርዝርዎ ለግምገማ ቀርቧል (ጥያቄ መለያ: %s)። ሲጸድቅ የመቤዠያ ኮድዎን እዚሁ ይደርስዎታል።',
        'already_submitted': 'ግዢዎ በአስተዳዳሪው እየተገመገመ ነው። ሁኔታዎን ለማየት /myaccess ይላኩ።',
        'already_approved': 'ቀደም ሲል የጸደቀ ግዢ አለዎት (ጥያቄ መለያ: %s)። የመቤዠያ ኮድዎን ለማየት /myaccess ይላኩ።',
        'myaccess_redeemed': 'የመቤዠያ ኮድዎ በ %s ተወስዷል።',
        'myaccess_voucher': 'ለ %s የመቤዠያ ኮድዎ:\nኮድ: %s',
        'no_voucher': 'ምንም ንቁ የመቤዠያ ኮድ አልተገኘም። ፓኬጅ ለመምረጥ እና ለመመዝገብ /start ይላኩ።',
        'cancelled': 'ሂደቱ ተሰርዟል። እንደገና ለመጀመር /start ይላኩ።',
        'approved_msg': '🎉 *ለ %s ያደረጉት ክፍያ ጸድቋል!*\n\n📱 *ስልክ ቁጥር:*\n`%s`\n\n🔑 *የመቤዠያ ኮድ:*\n`%s`\n\nለመጠቀም፡ የተምህሮ መተግበሪያን በመክፈት ስልክ ቁጥርዎን እና የመቤዠያ ኮድዎን ያስገቡ።',
        'copy_phone': '📱 %s',
        'copy_code': '🔑 %s',
        'btn_copy_phone': '📋 ስልክ ቁጥር ኮፒ ያድርጉ',
        'btn_copy_code': '🔑 የመቤዠያ ኮድ ኮፒ ያድርጉ',
        'rejected_msg': '❌ ለ %s ያቀረቡት የክፍያ ጥያቄ በአስተዳዳሪው ውድቅ ተደርጓል (ጥያቄ መለያ: %s)። ስህተት ነው ብለው ካመኑ እባክዎ ድጋፍ ሰጪውን ያነጋግሩ።',
        'approved_pending_pool': 'ለ %s ያደረጉት ክፍያ ጸድቋል! የድጋፍ ቡድናችን የመቤዠያ ኮድዎን በቅርቡ ይልካል (ጥያቄ መለያ: %s)።',
    },
    'om': {
        'lang_select': 'Maaloo afaan filattan ይምረጡ:',
        'start': 'Baga gara Application Barnoota Temhiro nagaan dhuftan! Maaloo itti fufuuf paakeejii filadhaa:',
        'package_chosen': 'Paakeejii: %s\nGatiisaa: %s\nMallaattoo gadii fayyadamuun lakkoofsa bilbila keessani qoodaa:',
        'share_phone_btn': '📱 Lakkoofsa Bilbilaa Qoodaa',
        'invalid_phone': 'Maaloo mallaattoo seeraa "Lakkoofsa Bilbilaa Qoodaa" jedhu gadii fayyadamaa.',
        'payment_method': 'Akkaamitti kafaluu meedhu? Maaloo mala kafaltii keessani filadhaa:',
        'invalid_method': 'Mali kafaltii dogoggoraa. Maaloo mallaattoowwan gadii fayyadamuun filadhaa.',
        'ask_name': 'Baye\'ee gaarii. Maqaa keessani guutuu galchaa:',
        'ask_link': 'Linkii risiitii kafaltii keessinii guutuu ergaa.\nBifa: %s',
        'invalid_link': 'Linkiin sun risiitii %s sirrii hin fakkaatu.\nBifa eegamu: %s\nMaaloo linkii guutuu ammas ergaa.',
        'ask_txid': 'Aamma lakkoofsa kaffaltii (Transaction ID / Ref Number) ergaa.\nBifa: %s',
        'invalid_txid': 'Lakkoofsi kaffaltii sun %s kuf sirrii hin fakkaatu.\nBifa eegamu: %s\nMaaloo ammas yaalaa.',
        'submitted_pending': 'Galatoomaa! Odeeffannoon kafaltii keessanii madaalliif dhihaateera (ID Gaaffii: %s). Yeroo mirkanaa\'u koodii fayyadamaa asitti ni argattu.',
        'already_submitted': 'Bitani keessani amma gulaaltuun irratti hojjechaa jira. Haala keessan ilaaluuf /myaccess ergaa.',
        'already_approved': 'Kafaltii mirkanaa\'e qabdu (ID Gaaffii: %s). Koodii keessan ilaaluuf /myaccess ergaa.',
        'myaccess_redeemed': 'Koodiin keessani Guyyaa %s irratti fudhatameera.',
        'myaccess_voucher': 'Koodii kaffaltii %s keessan:\nJecha Koodii: %s',
        'no_voucher': 'Koodiin hojii irratti jiru hin argamne. Paakeejii filachuufi galmaa\'uuf /start ergaa.',
        'cancelled': 'Adeemsi dhiikfameera. Irra deebitanii jalqabuuf /start ergaa.',
        'approved_msg': '🎉 *Kafaltiin keessan %s faana mirkanaa\'eera!*\n\n📱 *Lakkoofsa Bilbilaa:*\n`%s`\n\n🔑 *Koodii Fayyadamaa:*\n`%s`\n\nFayyadamuuf: Appii Temhiro Banuudhaan Lakkoofsa Bilbilaa fi Koodii keessani galchaa.',
        'copy_phone': '📱 %s',
        'copy_code': '🔑 %s',
        'btn_copy_phone': '📋 Lakkoofsa Bilbilaa Waraabi',
        'btn_copy_code': '🔑 Koodii Fayyadamaa Waraabi',
        'rejected_msg': '❌ Kaffaltiin %s irraatti dhihaate gulaalaan kufaa ta\'eera (ID Gaaffii: %s). Dogoggora jadhanii yoo yaaddu deeggarsa qunnamaa.',
        'approved_pending_pool': 'Kafaltiin keessan %s faana mirkanaa\'eera! Gareen deeggarsaa keenya dhiyootti koodii ergaa (ID Gaaffii: %s).',
    },
    'so': {
        'lang_select': 'Fadlan dooro luqadda aad doorbideyso:',
        'start': 'Ku soo dhowaw App-ka Waxbarashada ee Temhiro! Fadlan dooro xزمة (package) si aad u jarayso:',
        'package_chosen': 'Xزمة: %s\nQiimaha: %s\nFadlan la wadaag nambarkaaga telefonka adigoo kullanaya batoonka hoose:',
        'share_phone_btn': '📱 La Wadaag Nambarka Telefonka',
        'invalid_phone': 'Fadlan isticmaal batoonka rasmiga ah ee "La Wadaag Nambarka Telefonka" ee hoose.',
        'payment_method': 'Sidaad u rabtaa inaad u bixiso? Fadlan dooro habka lacag bixinta:',
        'invalid_method': 'Habka lacag bixinta ee saxda ahayn. Fadlan dooro adigoo isticmaalaya batoonada hoose.',
        'ask_name': 'Aad uga wanaagsan. Waa maxay magacaaga buuxa?',
        'ask_link': 'Soo dir xiriirka (link) rasiidka lacag bixinta buuxa.\nHabka: %s',
        'invalid_link': 'Xiriirkaasi ma u muuqdo rasiidh %s oo sax ah.\nHabka la filayo: %s\nFadlan soo dir xiriirka buuxa markale.',
        'ask_txid': 'Hada soo dir ID-ga maamulka / Nambarka tixraaca (Ref Number).\nHabka: %s',
        'invalid_txid': 'ID-ga maamulkaasi ma u muuqdo mid sax ah %s.\nHabka la filayo: %s\nFadlan dib u doondoon marka kale.',
        'submitted_pending': 'Waad mahadsan tahay! Faahfaahinta lacag bixintaada waxaa loo gudbiyay dib u eegis (ID Request: %s). Waxaad heli doontaa koodka voucher-kaaga halkan marka la ansixiyo.',
        'already_submitted': 'Raskaga bixinta wuxuu hadda ku jiraa dib u eegis maamule. Soo dir /myaccess si aad u eegto xaaladdaada.',
        'already_approved': 'Waxaad horay u leedahay rasiidh la ansixiyay (ID Request: %s). Soo dir /myaccess si aad u eegto koodka voucher-kaaga.',
        'myaccess_redeemed': 'Voucher-kaaga waxaa la furay %s.',
        'myaccess_voucher': 'Voucher-kaaga %s:\nErayga: %s',
        'no_voucher': 'Laguma helin voucher firfircoon. Isticmaal /start si aad u doorato xزمة una isdhaafiso.',
        'cancelled': 'Hawshii waa la baabi\'iyay. Soo dir /start si aad dib ugu start gareyso.',
        'approved_msg': '🎉 *Lacag bixintaada ee %s waa la ansixiyay!*\n\n📱 *Nambarka Telefonka:*\n`%s`\n\n🔑 *Koodka Voucher-ka:*\n`%s`\n\nSi aad u furato, fur Temhiro Study App oo geli nambarkaaga telefonka iyo koodka voucher-ka.',
        'copy_phone': '📱 %s',
        'copy_code': '🔑 %s',
        'btn_copy_phone': '📋 Koobi Garayso Nambarka Telefonka',
        'btn_copy_code': '🔑 Koobi Garayso Koodka Voucher-ka',
        'rejected_msg': '❌ Codsigaaga lacag bixinta ee %s waxaa soo diiday maamulaha (ID Request: %s). Haddii aad u meel dhagaxdo in tani tahay qalad, fadlan la xiriir taageerada.',
        'approved_pending_pool': 'Lacag bixintaada ee %s waa la ansixiyay! Kooxda taageerada waxay kuu soo diri doonaan koodka voucher-ka dhowaan (ID Request: %s).',
    },
    'ti': {
        'lang_select': 'በጃኹም ዝመርጽዎ ቋንቋ ይረዩ:',
        'start': 'እንካዕ ናብ ናይ ተምህሮ መጽናዕቲ መተግበሪ ብደሓን መጻእኹም! ንምቕጻል በጃኹም ፓኬጅ ይረዩ:',
        'package_chosen': 'ፓኬጅ: %s\nዋጋ: %s\nበጃኹም ኣብ ታሕቲ ዘሎ መላገቢ ብምጥቃም ቁጽሪ ስልኽኩም ኣካፍሉ:',
        'share_phone_btn': '📱 ቁጽሪ ስልኺ ኣካፍል',
        'invalid_phone': 'በጃኹም ኣብ ታሕቲ ዘሎ ወግዓዊ "ቁጽሪ ስልኺ ኣካፍል" መላገቢ ይጠቀሙ።',
        'payment_method': 'ብኸመይ ክትከፍሉ ትደልዩ? በጃኹም ናይ ክፍሊት መንገዲ ይረዩ:',
        'invalid_method': 'ጌጋ ናይ ክፍሊት መንገዲ። በጃኹም ኣብ ታሕቲ ዘለው መላገቢታት ብምጥቃም ይረዩ።',
        'ask_name': 'ብጣዕሚ ጽቡቕ። ምሉእ ስምኩም የእትዉ:',
        'ask_link': 'ምሉእ ናይ ክፍሊትኩም ደረሰኝ ሊንክ ይላኹ።\nቅዲ: %s',
        'invalid_link': 'እቲ ዝላኣኽክምዎ ሊንክ ናይ %s ደረሰኝ ኣይመስልን።\nዝድለ ቅዲ: %s\nበጃኹም ምሉእ ሊንክ እንደገና ይላኹ።',
        'ask_txid': 'ሕዚ ናይ መለለዪ ቁጽሪ ክፍሊትኩም (Transaction ID / Ref Number) ይላኹ።\nቅዲ: %s',
        'invalid_txid': 'እቲ ዘእተወክምዎ ናይ መለለዪ ቁጽሪ ክፍሊት ን %s ትክክለኛ ኣይመስልን።\nዝድለ ቅዲ: %s\nበጃኹም እንደገና ይሞክሩ።',
        'submitted_pending': 'የቐንየልና! ናይ ክፍሊት ዝርዝርኩም ንግምገማ ቐሪቡ ኣሎ (ጥያቄ መለያ: %s)። ምስ ጸደቐ ናይ መቤዠያ ኮድኩም ኣብዚ ክበጽሓኩም እዩ።',
        'already_submitted': 'ዓደግቲክም ብኣካየዲ ይግምገም ኣሎ። ኩነታትኩም ንምርኣይ /myaccess ይላኹ።',
        'already_approved': 'ኣቐዲሙ ዝጸደቐ ዓደግቲ ኣለኩም (ጥያቄ መለያ: %s)። ናይ መቤዠያ ኮድኩም ንምርኣይ /myaccess ይላኹ።',
        'myaccess_redeemed': 'ናይ መቤዠያ ኮድኩም ኣብ %s ተወሲዱ እዩ።',
        'myaccess_voucher': 'ናይ %s መቤዠያ ኮድኩም:\nኮድ: %s',
        'no_voucher': 'ዝኾነ ንጡፍ ናይ መቤዠያ ኮድ ኣይተረኸበን። ፓኬጅ ንምምራፅን ንምምዝጋብን /start ይላኹ።',
        'cancelled': 'እቲ መስርሕ ተሰሪዙ እዩ። እንደገና ንምጅማር /start ይላኹ።',
        'approved_msg': '🎉 *ናይ %s ክፍሊትኩም ጸድቑ ኣሎ!*\n\n📱 *ቁጽሪ ስልኺ:*\n`%s`\n\n🔑 *ናይ መቤዠያ ኮድ:*\n`%s`\n\nንምጥቃም፡ ናይ ተምህሮ መተግበሪ ብምክፋት ቁጽሪ ስልኽኩምን ናይ መቤዠያ ኮድኩምን የእትዉ።',
        'copy_phone': '📱 %s',
        'copy_code': '🔑 %s',
        'btn_copy_phone': '📋 ቁጽሪ ስልኺ ኮፒ ግበሩ',
        'btn_copy_code': '🔑 ናይ መቤዠያ ኮድ ኮፒ ግበሩ',
        'rejected_msg': '❌ ን %s ዝኣቐረብክምዎ ናይ ክፍሊት ሕቶ ብኣካየዲ ውድቂ ኾይኑ ኣሎ (ጥያቄ መለያ: %s)። ጌጋ እዩ ኢልኩም እንተኣሚንኩም በጃኹም ናይ ደጋፊ ኣገልግሎት ኣነጋግሩ።',
        'approved_pending_pool': 'ናይ %s ክፍሊትኩም ጸድቑ ኣሎ! ናይ ደጋፊ ቡድና ኣብ ቐረባ እዋን ናይ መቤዠያ ኮድኩም ክልእኽ እዩ (ጥያቄ መለያ: %s)።',
    },
}

def get_user_lang(user_id):
    data = read_json(USERS_FILE, default_users())
    user_info = (data.get('users') or {}).get(str(user_id)) or {}
    return user_info.get('language') or 'en'

def set_user_lang(user_id, lang_code):
    def _mutate(data):
        users = data.setdefault('users', {})
        user_entry = users.setdefault(str(user_id), {'id': user_id})
        user_entry['language'] = lang_code
        return True
    mutate_json(USERS_FILE, default_users(), _mutate)

def get_msg(user_id, key, *args):
    lang = get_user_lang(user_id)
    text = MESSAGES.get(lang, MESSAGES['en']).get(key) or MESSAGES['en'].get(key, '')
    if args:
        try:
            return text % args
        except Exception:
            return text
    return text

def build_language_keyboard():
    rows = []
    keys = list(LANGUAGES.keys())
    for i in range(0, len(keys), 2):
        row = [{'text': LANGUAGES[k], 'callback_data': 'lang:%s' % k} for k in keys[i:i+2]]
        rows.append(row)
    return {'inline_keyboard': rows}

def build_category_keyboard():
    rows = [
        [{'text': '🎓 Freshman', 'callback_data': 'cat:freshman'}],
        [{'text': '🏛 University Department', 'callback_data': 'cat:university'}],
        [{'text': '📋 COC Exam Preparation', 'callback_data': 'cat:coc'}],
        [{'text': '📚 EUEE / UAT / Other Packages', 'callback_data': 'cat:other'}]
    ]
    return {'inline_keyboard': rows}

def build_freshman_stream_keyboard():
    rows = [
        [{'text': '🧪 Natural Science (Year 1)', 'callback_data': 'freshman:nat'}],
        [{'text': '📐 Social Science (Year 1)', 'callback_data': 'freshman:soc'}],
        [{'text': '🔙 Back to Categories', 'callback_data': 'cat:main'}]
    ]
    return {'inline_keyboard': rows}

def build_plan_keyboard(prefix):
    rows = [
        [{'text': 'Semester 1 — 300 ETB', 'callback_data': '%s_sem1' % prefix}],
        [{'text': 'Semester 2 — 300 ETB', 'callback_data': '%s_sem2' % prefix}],
        [{'text': '🌟 Full Academic Year — 500 ETB (Unlocks Both)', 'callback_data': '%s_full_year' % prefix}],
        [{'text': '🔙 Back', 'callback_data': 'cat:main'}]
    ]
    return {'inline_keyboard': rows}

def build_university_year_keyboard():
    rows = [
        [{'text': 'Year 2', 'callback_data': 'univ_yr:2'}, {'text': 'Year 3', 'callback_data': 'univ_yr:3'}],
        [{'text': 'Year 4', 'callback_data': 'univ_yr:4'}, {'text': 'Year 5', 'callback_data': 'univ_yr:5'}],
        [{'text': '🔙 Back to Categories', 'callback_data': 'cat:main'}]
    ]
    return {'inline_keyboard': rows}

def build_university_dept_keyboard(year):
    departments = [
        ('📊 Accounting & Finance', 'accounting_finance'),
        ('📈 Economics', 'economics'),
        ('👔 Management', 'management'),
        ('📢 Marketing Management', 'marketing_management'),
        ('📦 LSCM (Logistics & Supply Chain)', 'lscm'),
        ('💼 BAIS (Business Admin & Info Sys)', 'bais'),
        ('🏛 PADM (Public Admin & Dev Mgt)', 'padm'),
        ('💻 Computer Science', 'computer_science'),
        ('💻 Software Engineering', 'software_engineering'),
        ('ℹ️ Information Sciences', 'information_sciences'),
        ('⚡ Electrical Engineering', 'electrical_engineering'),
        ('⚙️ Mechanical Engineering', 'mechanical_engineering'),
        ('🧠 Psychology', 'psychology'),
        ('🌐 PSIR (Political Sci & Int Rel)', 'psir'),
        ('⚖️ Ethiopian Law', 'ethiopian_law'),
    ]
    rows = []
    for label, dept_key in departments:
        rows.append([{'text': label, 'callback_data': 'univ_dept:%s:y%s' % (dept_key, year)}])
    rows.append([{'text': '🔙 Back to Years', 'callback_data': 'cat:university'}])
    return {'inline_keyboard': rows}

def build_coc_keyboard():
    departments = [
        ('🩺 Medicine COC Exam', 'coc_medicine'),
        ('🦷 Dental Medicine COC Exam', 'coc_dentistry'),
        ('💊 Pharmacy COC Exam', 'coc_pharmacy'),
        ('🏗 Engineering COC Exam', 'coc_engineering'),
        ('🏛 Architecture COC Exam', 'coc_architecture'),
        ('💻 Computer Science & Info IT/Science COC', 'coc_cs_is'),
        ('⚖️ Law COC Exam', 'coc_law'),
    ]
    rows = [[{'text': '%s — 300 ETB' % label, 'callback_data': 'package:%s' % key}] for label, key in departments]
    rows.append([{'text': '🔙 Back to Categories', 'callback_data': 'cat:main'}])
    return {'inline_keyboard': rows}

def build_package_keyboard():
    rows = [[
        {'text': '%s - %s' % (pkg['label'], format_money(pkg['priceCents'], pkg['currency'])),
         'callback_data': 'package:%s' % pkg['key']}
    ] for pkg in load_packages()]
    return {'inline_keyboard': rows}

def build_phone_share_keyboard(user_id=None):
    btn_text = get_msg(user_id, 'share_phone_btn') if user_id else '📱 Share Phone Number'
    return {
        'keyboard': [[{'text': btn_text, 'request_contact': True}]],
        'resize_keyboard': True,
        'one_time_keyboard': True,
    }

def get_start_message(user_id=None):
    return get_msg(user_id, 'start') if user_id else MESSAGES['en']['start']

# ---------------------------------------------------------------------------
# Message Builders
# ---------------------------------------------------------------------------

def build_request_message(request):
    user = request.get('user') or {}
    phone = request.get('phone') or {}
    phone_text = phone.get('number') or 'n/a'
    phone_verified = ' (verified)' if phone.get('verified') else ' (unverified)'
    full_name = ('%s %s' % (user.get('firstName') or '', user.get('lastName') or '')).strip()
    ref_line = ('Purchase Ref: `%s`\n' % request.get('purchaseReference')) if request.get('purchaseReference') else ''
    return '\n'.join([
        '📌 *New Access Request Pending Review*',
        'Request ID: `%s`' % request.get('requestId'),
        '%sProduct/Package: %s (%s)' % (ref_line, request.get('packageLabel'), format_money(request.get('priceCents'), request.get('currency'))),
        'Payment Method: %s' % (request.get('paymentMethodLabel') or request.get('paymentMethod') or 'N/A'),
        'Payer Name: %s' % (request.get('name') or full_name),
        'Telegram User: @%s (`%s`)' % (user.get('username') or 'no_username', user.get('id')),
        'Phone: `%s`%s' % (phone_text, phone_verified),
        'Transaction ID: `%s`' % request.get('transactionId'),
        'Transaction Link: %s' % request.get('transactionLink'),
        '',
        'Approve this request only after verifying the payment.',
    ])

def build_pending_message(request):
    user_id = request.get('userId')
    return get_msg(user_id, 'submitted_pending', request.get('requestId'))

def build_approved_pending_voucher_message(request):
    user_id = request.get('userId')
    return get_msg(user_id, 'approved_pending_pool', request.get('packageLabel'), request.get('requestId'))

def build_rejected_message(request):
    user_id = request.get('userId')
    return get_msg(user_id, 'rejected_msg', request.get('packageLabel'), request.get('requestId'))

def build_approved_message(package_label, voucher_phrase, user_id=None, phone_number=None):
    text = get_msg(user_id, 'approved_msg', package_label, phone_number or 'N/A', voucher_phrase)
    keyboard = {
        'inline_keyboard': [
            [{'text': get_msg(user_id, 'btn_copy_phone'), 'callback_data': 'copy_phone:%s' % (phone_number or '')}],
            [{'text': get_msg(user_id, 'btn_copy_code'), 'callback_data': 'copy_code:%s' % voucher_phrase}],
        ]
    }
    return text, keyboard

def build_pending_assignment_message(pending):
    lines = ['Approved purchases waiting for voucher assignment:']
    if not pending:
        lines.append('(none)')
    for index, req in enumerate(pending, start=1):
        user = req.get('user') or {}
        phone = (req.get('phone') or {}).get('number') or 'n/a'
        verified = ' (verified)' if (req.get('phone') or {}).get('verified') else ' (unverified)'
        full_name = ('%s %s' % (user.get('firstName') or '', user.get('lastName') or '')).strip()
        lines.append('%d. %s | %s | %s | @%s | %s%s'
                     % (index, req.get('requestId'), req.get('packageLabel'), full_name,
                        user.get('username') or 'no_username', phone, verified))
    return '\n'.join(lines)

def notify_admin(request):
    if not ADMIN_CHAT_ID:
        raise RuntimeError('Missing ADMIN_CHAT_ID in environment')
    send_message(ADMIN_CHAT_ID, build_request_message(request), parse_mode='Markdown', reply_markup={
        'inline_keyboard': [[
            {'text': '✅ Approve', 'callback_data': 'approve:%s' % request['requestId']},
            {'text': '❌ Reject', 'callback_data': 'reject:%s' % request['requestId']},
        ]]
    })


# ---------------------------------------------------------------------------
# Voucher pools
# ---------------------------------------------------------------------------


class PoolEmptyError(Exception):
    pass


def _load_pool_phrases(pool_key):
    """Read phrases for a pool from a text file inside PHRASES_DIR.

    File layout:
        <PHRASES_DIR>/<pool_key>.txt        -> phrase per line, ignores blanks
        <PHRASES_DIR>/<pool_key>@<suffix>.txt -> lines starting with '<phrase>:'
                                               are split into phrase / extra
    """
    directory = PHRASES_DIR
    phrase_file = os.path.join(directory, '%s.txt' % pool_key)
    if os.path.exists(phrase_file):
        with open(phrase_file, 'r', encoding='utf-8') as fh:
            lines = [line.strip() for line in fh if line.strip()]
        return [line for line in lines if not line.startswith('#')]

    entries = []
    if os.path.isdir(directory):
        prefix = '%s@' % pool_key
        for name in os.listdir(directory):
            if not name.startswith(prefix) or not name.endswith('.txt'):
                continue
            with open(os.path.join(directory, name), 'r', encoding='utf-8') as fh:
                for line in fh:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if ':' in line:
                        phrase, _, extra = line.partition(':')
                        entries.append({'phrase': phrase.strip(), 'extra': extra.strip()})
                    else:
                        entries.append({'phrase': line, 'extra': ''})
    return entries


def _as_pool_entry(item):
    if isinstance(item, dict):
        return {'phrase': str(item.get('phrase') or item.get('code') or ''), 'extra': str(item.get('extra') or '')}
    return {'phrase': str(item), 'extra': ''}


def _normalize_pool(pool):
    """Migrate older pool layouts (available/issued, plain-string lists) into
    the current unused/assigned dict layout so reserved phrases are never
    double-issued when old data is reused."""
    if 'unused' not in pool:
        pool['unused'] = []
        old_available = pool.pop('available', None)
        if old_available:
            pool['unused'] = [_as_pool_entry(item) for item in old_available]
    if 'assigned' not in pool:
        pool['assigned'] = []
        old_issued = pool.pop('issued', None)
        if isinstance(old_issued, dict):
            pool['assigned'] = [_as_pool_entry(item) for item in old_issued.values()]
        elif isinstance(old_issued, list):
            pool['assigned'] = [_as_pool_entry(item) for item in old_issued]
    return pool


def _hydrate_pool(data, pool_key):
    """Ensure 'data' has a complete entry for pool_key. Loads every phrase that
    is not already tracked in the store so new phrase files show up in
    /assign without a restart."""
    pool = data['packages'].get(pool_key)
    if pool is None:
        pool = {'unused': [], 'assigned': []}
        data['packages'][pool_key] = pool
    pool = _normalize_pool(pool)

    unused_by_phrase = {str(item['phrase']) for item in pool['unused']}
    assigned_by_phrase = {str(item['phrase']) for item in pool['assigned']}

    loaded = _load_pool_phrases(pool_key)
    for entry in loaded:
        phrase = entry if isinstance(entry, str) else entry.get('phrase')
        extra = '' if isinstance(entry, str) else entry.get('extra', '')
        key = str(phrase)
        if key in unused_by_phrase or key in assigned_by_phrase:
            continue
        pool['unused'].append({'phrase': key, 'extra': extra})

    return pool


def reserve_voucher(pool_key):
    """Pick and remove the first unused voucher for a pool atomically.
    Returns {'phrase': ..., 'extra': ...} or raises PoolEmptyError."""
    def _reserve(data):
        pool = _hydrate_pool(data, pool_key)
        if not pool['unused']:
            raise PoolEmptyError('Pool "%s" is empty' % pool_key)
        entry = pool['unused'].pop(0)
        assigned = pool.setdefault('assigned', [])
        assigned.append(dict(entry))
        return dict(entry)

    return mutate_json(VOUCHERS_FILE, default_vouchers(), _reserve)


def record_voucher_assignment(entry):
    """Persist a fully-assigned voucher (with owner details) into the 'issued'
    map so get_voucher_for_user / find_voucher / redeem_code can find it."""
    def _record(data):
        issued = data.setdefault('issued', {})
        issued[str(entry.get('phrase'))] = entry
        return True

    return mutate_json(VOUCHERS_FILE, default_vouchers(), _record)


def get_voucher_for_user(user_id):
    """Look up the most recently assigned, not-yet-redeemed voucher of a user."""
    data = read_json(VOUCHERS_FILE, default_vouchers())
    issued = data.get('issued') or {}
    entries = []
    for entry in issued.values():
        if str(entry.get('ownerId')) != str(user_id):
            continue
        status = normalize_voucher_status(entry)
        if status in ('assigned', 'issued'):
            entries.append(entry)
    if not entries:
        return None
    entries.sort(key=lambda item: str(item.get('assignedAt') or ''), reverse=True)
    return entries[0]


def normalize_voucher_status(entry):
    if entry.get('redeemed'):
        return 'redeemed'
    if entry.get('disbursed') or entry.get('sent'):
        return 'disbursed'
    if entry.get('assigned'):
        return 'assigned'
    return 'issued'


# ---------------------------------------------------------------------------
# Redemption API helpers
# ---------------------------------------------------------------------------


def verify_init_data(init_data):
    """Validate a Telegram Mini App initData string.

    Returns the parsed user object (dict) when the data is authentic and
    within INIT_DATA_MAX_AGE_SECONDS, otherwise returns None.

    initData format (urlencoded pairs, sorted, joined with '\\n'):
        auth_date=<unix>\\nquery_id=<..>\\nuser=<urlencoded-json>...&hash=<sha256>
    The hash is the HMAC-SHA256 of that signature string using a secret key
    derived from the bot token.
    """
    if not init_data or not TELEGRAM_BOT_TOKEN:
        return None

    try:
        parsed = dict(pair.split('=', 1) for pair in str(init_data).split('&'))
    except ValueError:
        return None

    received_hash = parsed.pop('hash', None)
    if not received_hash:
        return None

    auth_date = parsed.get('auth_date')
    try:
        if abs(int(auth_date) - time.time()) > INIT_DATA_MAX_AGE_SECONDS:
            return None
    except (TypeError, ValueError):
        return None

    signature = '\n'.join('%s=%s' % (k, parsed[k]) for k in sorted(parsed))
    secret_key = hmac.new(b'WebAppData', TELEGRAM_BOT_TOKEN.encode('utf-8'), hashlib.sha256).digest()
    expected_hash = hmac.new(secret_key, signature.encode('utf-8'), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected_hash, received_hash):
        return None

    try:
        user = json.loads(unquote(parsed.get('user', '{}')))
    except ValueError:
        user = None
    return user if isinstance(user, dict) else None


def default_auth_tokens():
    return {'tokens': {}}


def verify_login_widget(data):
    """Validate a Telegram Login Widget / native-login payload.

    The native Android "Login with Telegram" SDK (telegram-login-android)
    returns a signed set of fields: id, first_name, last_name, username,
    photo_url, auth_date and hash. The hash is an HMAC-SHA256 of the
    sorted key=value pairs using the SHA-256 of the bot token as the key
    (the same scheme as the web Login Widget). Returns the user dict or
    None when the data is not authentic or is stale.
    """
    if not data or not TELEGRAM_BOT_TOKEN:
        return None

    try:
        payload = dict(data)
    except (TypeError, ValueError):
        return None

    received_hash = payload.pop('hash', None)
    if not received_hash:
        return None

    auth_date = payload.get('auth_date')
    try:
        if abs(int(auth_date) - time.time()) > INIT_DATA_MAX_AGE_SECONDS:
            return None
    except (TypeError, ValueError):
        return None

    data_check_string = '\n'.join(
        '%s=%s' % (k, payload[k]) for k in sorted(payload)
    )
    secret_key = hashlib.sha256(TELEGRAM_BOT_TOKEN.encode('utf-8')).digest()
    expected_hash = hmac.new(secret_key, data_check_string.encode('utf-8'), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected_hash, received_hash):
        return None

    user_id = payload.get('id')
    if user_id is None:
        return None

    return {
        'id': str(user_id),
        'username': str(payload.get('username') or ''),
        'firstName': str(payload.get('first_name') or ''),
        'lastName': str(payload.get('last_name') or ''),
    }


def issue_auth_token(user_id):
    """Create a short-lived token bound to a Telegram user id."""
    token = secrets.token_urlsafe(32)

    def _store(data):
        data.setdefault('tokens', {})[token] = {
            'userId': str(user_id),
            'expiresAt': utcnow(offset_seconds=AUTH_TOKEN_TTL_SECONDS),
        }
        return token

    mutate_json(AUTH_TOKENS_FILE, default_auth_tokens(), _store)
    return token


def resolve_auth_token(token):
    """Map an auth token back to its Telegram user id, or None if unknown
    or expired."""
    if not token:
        return None
    token = str(token)

    def _read(data):
        entry = (data.get('tokens') or {}).get(token)
        if not entry:
            return None
        try:
            expires_at = datetime.fromisoformat(str(entry.get('expiresAt')).replace('Z', '+00:00'))
        except ValueError:
            return None
        if expires_at < datetime.now(timezone.utc):
            return None
        return str(entry.get('userId'))

    return mutate_json(AUTH_TOKENS_FILE, default_auth_tokens(), _read)


def default_access_tokens():
    return {'tokens': {}}


def issue_access_token(entitlement):
    """Issue a long-lived app access token bound to an entitlement."""
    token = secrets.token_urlsafe(32)
    now = utcnow()

    def _store(data):
        data.setdefault('tokens', {})[token] = {
            'entitlementId': entitlement.get('entitlementId'),
            'phone': normalize_phone(entitlement.get('phone')),
            'userId': str(entitlement.get('ownerId') or ''),
            'packageKey': entitlement.get('packageKey'),
            'packageLabel': entitlement.get('packageLabel'),
            'issuedAt': now,
            'expiresAt': utcnow(offset_seconds=ACCESS_TOKEN_TTL_SECONDS),
            'revoked': False,
        }
        return token

    mutate_json(ACCESS_TOKENS_FILE, default_access_tokens(), _store)
    return token


def resolve_access_token(token):
    """Return (status, entry) for an access token.

    status is 'valid' | 'unknown' | 'revoked' | 'expired'."""
    token = str(token or '').strip()
    if not token:
        return 'unknown', None

    def _read(data):
        entry = (data.get('tokens') or {}).get(token)
        if not entry:
            return 'unknown', None
        if entry.get('revoked'):
            return 'revoked', entry
        try:
            expires_at = datetime.fromisoformat(str(entry.get('expiresAt')).replace('Z', '+00:00'))
        except ValueError:
            return 'expired', entry
        if expires_at < datetime.now(timezone.utc):
            return 'expired', entry
        return 'valid', entry

    return mutate_json(ACCESS_TOKENS_FILE, default_access_tokens(), _read)


def revoke_access_token(token):
    token = str(token or '').strip()
    if not token:
        return False

    def _revoke(data):
        entry = (data.get('tokens') or {}).get(token)
        if not entry:
            return False
        entry['revoked'] = True
        entry['revokedAt'] = utcnow()
        return True

    return mutate_json(ACCESS_TOKENS_FILE, default_access_tokens(), _revoke)


def revoke_entitlement(entitlement_id):
    entitlement_id = str(entitlement_id or '').strip()
    if not entitlement_id:
        return False

    def _revoke(data):
        entry = (data.get('entitlements') or {}).get(entitlement_id)
        if not entry:
            return False
        entry['revoked'] = True
        entry['revokedAt'] = utcnow()
        return True

    return mutate_json(ENTITLEMENTS_FILE, default_entitlements(), _revoke)


RATE_LIMIT = {}


def rate_limit_redeem(user_key, max_hits=REDEEM_RATE_LIMIT_MAX, window_ms=REDEEM_RATE_LIMIT_WINDOW_MS):
    now_ms = int(time.time() * 1000)
    hits = [ts for ts in RATE_LIMIT.get(user_key, []) if now_ms - ts < window_ms]
    if len(hits) >= max_hits:
        RATE_LIMIT[user_key] = hits
        return False
    hits.append(now_ms)
    RATE_LIMIT[user_key] = hits
    return True


def find_voucher(owner_id, phrase):
    data = read_json(VOUCHERS_FILE, default_vouchers())
    issued = data.get('issued') or {}
    for entry in issued.values():
        if str(entry.get('ownerId')) != str(owner_id):
            continue
        if str(entry.get('phrase')) == str(phrase):
            return entry
    return None


def redeem_code(owner_id, phrase, device_id=''):
    """Atomically mark an owned voucher as redeemed and write the entitlement.
    Returns the entitlement record or a dict error."""
    def _redeem(data):
        issued = data.setdefault('issued', {})
        for key, entry in issued.items():
            if str(entry.get('ownerId')) != str(owner_id):
                continue
            if str(entry.get('phrase')) != str(phrase):
                continue
            status = normalize_voucher_status(entry)
            if status == 'redeemed':
                return {'error': 'Voucher already redeemed'}
            if status != 'assigned':
                return {'error': 'Voucher is not assigned to you yet'}

            entry['redeemed'] = True
            entry['redeemedAt'] = utcnow()
            entry['status'] = 'redeemed'
            entry['redeemedByUserId'] = str(owner_id)
            entry['redeemedByDeviceId'] = str(device_id or '')

            entitlement = {
                'entitlementId': random_id(16),
                'ownerId': str(owner_id),
                'phrase': entry.get('phrase'),
                'packageKey': entry.get('packageKey'),
                'packageLabel': entry.get('packageLabel'),
                'assignedAt': entry.get('assignedAt'),
                'redeemedAt': entry['redeemedAt'],
                'phone': entry.get('phone'),
                'ownerName': entry.get('ownerName'),
                'deviceId': str(device_id or ''),
            }
            return entitlement

        return {'error': 'No matching voucher found'}

    result = mutate_json(VOUCHERS_FILE, default_vouchers(), _redeem)
    if isinstance(result, dict) and 'error' in result:
        return result

    entitlement = result

    def _record(data):
        data.setdefault('entitlements', {})[entitlement['entitlementId']] = entitlement
        return True

    mutate_json(ENTITLEMENTS_FILE, default_entitlements(), _record)
    return entitlement


def activate_voucher(phone, phrase):
    """Atomically activate a voucher by verified phone + redeem code.

    This is the Android-app activation path: identity is the phone number
    captured by Telegram's Share Contact (phoneVerified), not a Telegram ID.
    Every validation (voucher exists, assigned, not redeemed, not revoked,
    owner phone matches, phone was Telegram-verified, payment approved,
    package valid) runs inside the single serialized write lock on
    VOUCHERS_FILE, so two simultaneous activations of the same code can never
    both succeed.

    Returns the entitlement record, or a dict error."""
    phone_digits = normalize_phone(phone)
    phrase = str(phrase or '').strip()
    if not phone_digits or not is_plausible_phone(phone_digits):
        return {'error': 'Invalid phone number'}
    if not phrase:
        return {'error': 'Missing redeem code'}

    def _activate(data):
        issued = data.setdefault('issued', {})
        entry = issued.get(phrase)
        if entry is None:
            return {'error': 'No matching voucher found'}

        status = normalize_voucher_status(entry)
        if status == 'redeemed':
            return {'error': 'Voucher already redeemed'}
        if entry.get('revoked'):
            return {'error': 'Voucher was revoked'}
        if status != 'assigned':
            return {'error': 'Voucher is not assigned to you yet'}

        request = None
        if entry.get('requestId'):
            users = read_json(USERS_FILE, default_users())
            for candidate in (users.get('requests') or {}).values():
                if candidate.get('requestId') == entry.get('requestId'):
                    request = candidate
                    break

        stored_phone = normalize_phone(entry.get('phone'))
        if not stored_phone or stored_phone != phone_digits:
            return {'error': 'Phone does not match the voucher owner'}

        request_phone = (request or {}).get('phone') or {}
        phone_verified = bool(request_phone.get('verified'))
        if not phone_verified:
            phone_verified = bool(entry.get('phoneVerified'))
        if not phone_verified:
            return {'error': 'Phone number was not verified through Telegram'}

        if request is None or request.get('status') not in ('approved', 'delivered'):
            return {'error': 'Payment was not approved'}

        pkg = get_package_by_key(entry.get('packageKey'))
        if pkg is None:
            return {'error': 'Package is invalid'}

        entry['redeemed'] = True
        entry['redeemedAt'] = utcnow()
        entry['status'] = 'redeemed'
        entry['redeemedByPhone'] = phone_digits
        entry['activatedBy'] = 'android'

        pkg_key = entry.get('packageKey')
        entitlements = get_product_entitlements(pkg_key)
        entitlement = {
            'entitlementId': random_id(16),
            'ownerId': str(entry.get('ownerId') or (request or {}).get('userId') or ''),
            'phone': phone_digits,
            'phrase': entry.get('phrase'),
            'packageKey': pkg_key,
            'packageLabel': entry.get('packageLabel'),
            'entitlements': entitlements,
            'assignedAt': entry.get('assignedAt'),
            'redeemedAt': entry['redeemedAt'],
            'activatedBy': 'android',
            'requestId': entry.get('requestId'),
        }
        return entitlement

    result = mutate_json(VOUCHERS_FILE, default_vouchers(), _activate)
    if isinstance(result, dict) and 'error' in result:
        return result

    def _record(data):
        data.setdefault('entitlements', {})[result['entitlementId']] = result
        return True

    mutate_json(ENTITLEMENTS_FILE, default_entitlements(), _record)
    return result


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------


EXPORT_HEADERS = [
    'Request ID', 'Package', 'User', 'Telegram Username', 'User ID', 'Phone',
    'Payment Method', 'Transaction ID', 'Transaction Link', 'Voucher Phrase',
    'Status', 'Approved At', 'Voucher Assigned At', 'Voucher Owner ID',
    'Owner Phone', 'Owner Name', 'Price (ETB)',
]


def _price_etb(package_key):
    """Unit price in ETB for a package key, or None when unknown."""
    pkg = get_package_by_key(package_key) if package_key else None
    if pkg is None:
        return None
    try:
        return int(pkg.get('priceCents') or 0) / 100.0
    except (TypeError, ValueError):
        return None


def _sale_count(requests, approved_map=None):
    """Count paid requests (have a voucher or are approved/delivered)."""
    approved_map = approved_map or {}
    count = 0
    for req in requests:
        if approved_map.get(req.get('requestId')) is not None:
            count += 1
        elif req.get('status') in ('approved', 'delivered'):
            count += 1
    return count


def write_exports(requests_by_package):
    """Write one Excel file per package plus a totals file.

    requests_by_package: {package_key: [request, ...]}
    Voucher/assignment columns are filled from vouchers.json 'issued' entries.
    Returns a list of written file paths.
    """
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    voucher_data = read_json(VOUCHERS_FILE, default_vouchers())
    approved_map = {}
    for entry in (voucher_data.get('issued') or {}).values():
        if entry.get('requestId'):
            approved_map[entry['requestId']] = entry

    written = []
    for package_key, requests in requests_by_package.items():
        pkg = get_package_by_key(package_key)
        label = (pkg['label'] if pkg else package_key).replace(' ', '_')
        filename = '%s_%s.xlsx' % (datetime.now().strftime('%Y%m%d_%H%M%S'), label)
        file_path = os.path.join(EXPORTS_DIR, filename)
        write_excel_requests(file_path, requests, approved_map=approved_map)
        written.append(file_path)

    totals_path = os.path.join(EXPORTS_DIR,
                               '%s_all_packages.xlsx' % datetime.now().strftime('%Y%m%d_%H%M%S'))
    write_excel_totals(totals_path, requests_by_package)
    written.append(totals_path)
    return written


def write_excel_requests(file_path, requests, approved_map=None):
    """Rows: one per request. Voucher/assignment columns filled from the
    approved entries only when the request belongs to that package."""
    approved_map = approved_map or {}
    wb = Workbook()
    ws = wb.active
    ws.title = 'Requests'
    _write_excel_headers(ws, EXPORT_HEADERS)

    for req in requests:
        voucher = approved_map.get(req.get('requestId'))
        ws.append([
            req.get('requestId'),
            req.get('packageLabel'),
            ('%s %s' % ((req.get('user') or {}).get('firstName') or '',
                        (req.get('user') or {}).get('lastName') or '')).strip(),
            (req.get('user') or {}).get('username') or '',
            (req.get('user') or {}).get('id') or '',
            (req.get('phone') or {}).get('number') or '',
            req.get('paymentMethodLabel') or '',
            req.get('transactionId') or '',
            req.get('transactionLink') or '',
            (voucher or {}).get('phrase') if voucher else '',
            normalize_voucher_status(voucher) if voucher else req.get('status') or '',
            req.get('approvedAt') or '',
            (voucher or {}).get('assignedAt') if voucher else '',
            (voucher or {}).get('ownerId') if voucher else '',
            (voucher or {}).get('phone') if voucher else '',
            (voucher or {}).get('ownerName') if voucher else '',
            _price_etb(req.get('packageKey')) or '',
        ])

    _autofit_excel(ws)
    _write_total_sales_sheet(wb, requests, approved_map)
    wb.save(file_path)


def _write_total_sales_sheet(wb, requests, approved_map=None):
    """Add a 'Total Sales' tab that sums the price of each paid row."""
    package_key = next((r.get('packageKey') for r in requests if r.get('packageKey')), None)
    pkg = get_package_by_key(package_key) if package_key else None
    label = (pkg['label'] if pkg else package_key or '') or ''
    price = _price_etb(package_key)
    sales = _sale_count(requests, approved_map)

    ws = wb.create_sheet('Total Sales')
    _write_excel_headers(ws, ['Metric', 'Value'])
    ws.append(['Package', label])
    ws.append(['Unit price (ETB)', price if price is not None else ''])
    ws.append(['Sales count', sales])
    ws.append(['Total sales (ETB)', round(price * sales, 2) if price is not None else ''])
    _autofit_excel(ws)


def write_excel_totals(file_path, requests_by_package):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Summary'
    _write_excel_headers(ws, ['Package', 'Total Requests', 'Approved', 'Pending', 'Rejected'])

    for package_key, requests in requests_by_package.items():
        pkg = get_package_by_key(package_key)
        label = pkg['label'] if pkg else package_key
        approved = sum(1 for req in requests if req.get('status') == 'approved')
        pending = sum(1 for req in requests if req.get('status') in ('pending', 'processing'))
        rejected = sum(1 for req in requests if req.get('status') == 'rejected')
        ws.append([label, len(requests), approved, pending, rejected])

    _autofit_excel(ws)

    ts = wb.create_sheet('Total Sales')
    _write_excel_headers(ts, ['Package', 'Unit Price (ETB)', 'Sales Count', 'Total Sales (ETB)'])
    grand_total = 0.0
    for package_key, requests in requests_by_package.items():
        pkg = get_package_by_key(package_key)
        label = pkg['label'] if pkg else package_key
        price = _price_etb(package_key)
        sales = _sale_count(requests)
        row_total = round((price or 0.0) * sales, 2) if price is not None else ''
        grand_total += (price or 0.0) * sales if price is not None else 0.0
        ts.append([label, price if price is not None else '', sales, row_total])
    ts.append(['GRAND TOTAL', '', '', round(grand_total, 2)])
    _autofit_excel(ts)

    wb.save(file_path)


def _write_excel_headers(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)


def _autofit_excel(ws):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            value_len = len(str(cell.value))
            if value_len > widths.get(cell.column_letter, 0):
                widths[cell.column_letter] = value_len
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 60)


# ---------------------------------------------------------------------------
# Store / user / draft helpers
# ---------------------------------------------------------------------------


def _store_user_profile(data, user):
    uid = str(user.get('id'))
    users = data.setdefault('users', {})
    users[uid] = {
        'firstName': user.get('first_name') or user.get('firstName') or '',
        'lastName': user.get('last_name') or user.get('lastName') or '',
        'username': user.get('username') or '',
        'languageCode': user.get('language_code') or '',
        'lastSeen': utcnow(),
    }


def _get_user_profile(data, user_id):
    return (data.get('users') or {}).get(str(user_id)) or {}


def _new_request(user_id, profile, package, phone):
    now = utcnow()
    return {
        'requestId': random_id(12),
        'userId': str(user_id),
        'user': {
            'id': str(user_id),
            'firstName': profile.get('firstName') or '',
            'lastName': profile.get('lastName') or '',
            'username': profile.get('username') or '',
            'languageCode': profile.get('languageCode') or '',
        },
        'packageKey': package['key'],
        'packageLabel': package['label'],
        'priceCents': package['priceCents'],
        'currency': package['currency'],
        'phrasePool': package['phrasePool'],
        'phone': {
            'number': (phone or {}).get('number'),
            'verified': bool((phone or {}).get('verified')),
        },
        'status': 'draft',
        'paymentMethod': None,
        'paymentMethodLabel': None,
        'transactionLink': None,
        'transactionId': None,
        'createdAt': now,
        'updatedAt': now,
        'messageId': None,
        'approvedAt': None,
        'rejectedAt': None,
        'voucher': None,
    }


def _new_draft(user_id):
    return {
        'userId': str(user_id),
        'step': 'package',
        'package': None,
        'purchaseReference': None,
        'phone': None,
        'method': None,
        'name': None,
        'link': None,
        'txid': None,
        'updatedAt': utcnow(),
    }


def _get_draft(user_id):
    def _read(data):
        drafts = data.setdefault('drafts', {})
        return drafts.get(str(user_id))

    return mutate_json(USERS_FILE, default_users(), _read)


def _save_draft(user_id, draft):
    def _write(data):
        data.setdefault('drafts', {})[str(user_id)] = draft
        return True

    return mutate_json(USERS_FILE, default_users(), _write)


def _latest_request_for_user(requests, user_id):
    """Return the most recently created request of a user, or None."""
    matches = [req for req in requests.values() if str(req.get('userId')) == str(user_id)]
    if not matches:
        return None
    matches.sort(key=lambda req: str(req.get('createdAt') or ''), reverse=True)
    return matches[0]


def _submit(data, user_id, draft):
    """Create a new purchase request for a user from their finished draft."""
    requests = data.setdefault('requests', {})

    profile = _get_user_profile(data, user_id)
    package = get_package_by_key(draft.get('package')) or DEFAULT_PACKAGES[0]
    phone = draft.get('phone') or {}
    request = _new_request(user_id, profile, package, phone)
    request.update({
        'purchaseReference': draft.get('purchaseReference'),
        'paymentMethod': draft.get('method'),
        'paymentMethodLabel': (get_payment_method(draft.get('method')) or {}).get('label'),
        'name': draft.get('name'),
        'transactionLink': draft.get('link'),
        'transactionId': draft.get('txid'),
        'status': 'pending',
        'updatedAt': utcnow(),
    })
    requests[str(request['requestId'])] = request
    
    # Mark associated purchase request as submitted
    ref = draft.get('purchaseReference')
    if ref:
        def _update_p(p_data):
            p = (p_data.get('purchases') or {}).get(ref)
            if p:
                p['status'] = 'submitted'
                p['updatedAt'] = utcnow()
                p['telegramUserId'] = str(user_id)
            return True
        mutate_json(PURCHASES_FILE, default_purchases_store(), _update_p)

    return request


# ---------------------------------------------------------------------------
# Flow: state machine steps
# ---------------------------------------------------------------------------


def handle_draft_step(user, text):
    """Advance a buyer's draft by one step. Returns (draft, message, reply_markup)."""
    user_id = user.get('id')
    draft = _get_draft(user_id) or _new_draft(user_id)
    step = draft.get('step')
    message = None
    reply_markup = None
    clean_text = str(text or '').strip()

    # Purchase Reference Lookup check if text looks like TH-XXXXX
    if is_valid_purchase_reference_format(clean_text) and (step in ('package', 'reference') or clean_text.upper().startswith('TH-')):
        ref_data = get_purchase_by_reference(clean_text)
        if not ref_data:
            return draft, get_msg(user_id, 'invalid_ref'), None
        if ref_data.get('status') in ('redeemed', 'delivered'):
            return draft, get_msg(user_id, 'already_redeemed_ref'), None

        # Lock in package and reference
        product_id = ref_data.get('productId')
        package = get_package_by_key(product_id) or {
            'key': product_id,
            'label': ref_data.get('productLabel'),
            'priceCents': ref_data.get('priceCents'),
            'currency': ref_data.get('currency', 'ETB'),
            'phrasePool': ref_data.get('phrasePool', 'freshman'),
        }

        entitlements_str = '\n'.join(['✓ %s' % e.replace('_', ' ').title() for e in ref_data.get('entitlements', [product_id])])
        draft.update({
            'package': package['key'],
            'purchaseReference': ref_data['purchaseReference'],
            'step': 'ref_confirm',
            'updatedAt': utcnow(),
        })
        _save_draft(user_id, draft)

        message = get_msg(
            user_id, 'ref_summary',
            package['label'],
            format_money(package['priceCents'], package['currency']),
            entitlements_str
        )
        reply_markup = {
            'inline_keyboard': [[
                {'text': get_msg(user_id, 'btn_confirm_ref'), 'callback_data': 'confirm_ref:%s' % ref_data['purchaseReference']},
                {'text': get_msg(user_id, 'btn_cancel_ref'), 'callback_data': 'cancel_ref'},
            ]]
        }
        return draft, message, reply_markup

    if step == 'ref_confirm' and clean_text.startswith('confirm_ref:'):
        ref = clean_text.split(':', 1)[1]
        draft.update({'step': 'phone', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'package_chosen', draft.get('package'), format_money(get_product_price(draft.get('package'))))
        reply_markup = build_phone_share_keyboard(user_id)
        return draft, message, reply_markup

    if step == 'package':
        package = get_package_by_key(text)
        if not package:
            return draft, get_start_message(user_id), build_package_keyboard()
        draft.update({'package': package['key'], 'step': 'phone', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'package_chosen', package['label'], format_money(package['priceCents'], package['currency']))
        reply_markup = build_phone_share_keyboard(user_id)

    elif step == 'phone':
        draft.update({'phone': {'number': text, 'verified': False}, 'step': 'method', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'payment_method')
        reply_markup = build_payment_method_keyboard()

    elif step == 'method':
        method = get_payment_method(text)
        if not method:
            message = get_msg(user_id, 'invalid_method')
            reply_markup = build_payment_method_keyboard()
            return draft, message, reply_markup
        draft.update({'method': method['key'], 'step': 'name', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'ask_name')

    elif step == 'name':
        draft.update({'name': text.strip(), 'step': 'link', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'ask_link', link_format_hint(draft.get('method')))

    elif step == 'link':
        method = draft.get('method')
        if not is_valid_transaction_link(method, text):
            message = get_msg(user_id, 'invalid_link', get_payment_method(method)['label'], link_format_hint(method))
            return draft, message, None
        draft.update({'link': text.strip(), 'step': 'txid', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        message = get_msg(user_id, 'ask_txid', id_format_hint(method))

    elif step == 'txid':
        method = draft.get('method')
        if not is_valid_transaction_id(method, text):
            message = get_msg(user_id, 'invalid_txid', get_payment_method(method)['label'], id_format_hint(method))
            return draft, message, None
        draft.update({'txid': text.strip(), 'step': 'done', 'updatedAt': utcnow()})
        _save_draft(user_id, draft)
        return draft, None, None

    return draft, message, reply_markup


def build_already_approved_message(request):
    user_id = request.get('userId')
    return get_msg(user_id, 'already_approved', request.get('requestId'))


def submit_request(user_id, draft):
    """Submit a finished draft. Messages the buyer appropriately whether the
    request is new, already pending, already approved, or a re-submit after
    a rejection. Never goes silent."""
    def _do(data):
        request = _submit(data, user_id, draft)
        return request

    request = mutate_json(USERS_FILE, default_users(), _do)
    status = request.get('status')

    if status == 'pending':
        logger.info('New request %s (user %s) persisted; notifying admin', request.get('requestId'), user_id)
        send_message(user_id, build_pending_message(request))
        notify_admin(request)
    elif status in ('approved', 'pending_assignment', 'delivered'):
        send_message(user_id, build_already_approved_message(request))
    return request


def finalize_request(request_id, decision, admin_user):
    """decision: 'approved' or 'rejected'."""
    def _do(data):
        requests = data.get('requests') or {}
        for req in requests.values():
            if req.get('requestId') != request_id:
                continue
            if decision == 'approved':
                req['status'] = 'approved'
                req['approvedAt'] = utcnow()
                req['approvedBy'] = str(admin_user.get('id'))
            else:
                req['status'] = 'rejected'
                req['rejectedAt'] = utcnow()
                req['rejectedBy'] = str(admin_user.get('id'))
            req['updatedAt'] = utcnow()
            return req
        return None

    return mutate_json(USERS_FILE, default_users(), _do)


def assign_voucher(request_id, support_user):
    """Mark a pending_assignment request as delivered and reserve a voucher.
    Returns a status string: 'ok' | 'missing' | 'empty' | 'already'."""
    def _do(data):
        requests = data.get('requests') or {}
        req = None
        for candidate in requests.values():
            if candidate.get('requestId') == request_id:
                req = candidate
                break
        if req is None:
            return 'missing'

        if req.get('voucher'):
            return 'already'

        try:
            entry = reserve_voucher(req.get('phrasePool'))
        except PoolEmptyError:
            req['status'] = 'pending_assignment'
            req['updatedAt'] = utcnow()
            return 'empty'

        owner_profile = _get_user_profile(data, str(req.get('userId')))
        entry.update({
            'ownerId': str(req.get('userId')),
            'ownerName': ('%s %s' % (owner_profile.get('firstName') or '',
                                     owner_profile.get('lastName') or '')).strip(),
            'phone': (req.get('phone') or {}).get('number'),
            'phoneVerified': bool((req.get('phone') or {}).get('verified')),
            'packageKey': req.get('packageKey'),
            'packageLabel': req.get('packageLabel'),
            'assigned': True,
            'status': 'assigned',
            'assignedAt': utcnow(),
            'assignedBy': str(support_user.get('id')),
            'requestId': req.get('requestId'),
        })
        record_voucher_assignment(entry)
        req['voucher'] = {
            'phrase': entry.get('phrase'),
            'extra': entry.get('extra', ''),
            'assignedAt': entry.get('assignedAt'),
            'assignedBy': entry.get('assignedBy'),
            'ownerId': entry.get('ownerId'),
        }
        req['status'] = 'delivered'
        req['updatedAt'] = utcnow()
        return 'ok'

    return mutate_json(USERS_FILE, default_users(), _do)


def send_pending_summary(support_user):
    """Send the list of approved purchases that still need a voucher."""
    def _collect(data):
        pending = []
        for req in (data.get('requests') or {}).values():
            if req.get('status') in ('pending_assignment', 'approved'):
                pending.append(req)
        return pending

    pending = mutate_json(USERS_FILE, default_users(), _collect)
    keyboard = {'inline_keyboard': []}
    for req in pending:
        keyboard['inline_keyboard'].append([{
            'text': 'Assign: %s (%s)' % (req.get('requestId'), req.get('packageLabel')),
            'callback_data': 'assign:%s' % req.get('requestId'),
        }])
    if not keyboard['inline_keyboard']:
        keyboard = None
    send_message(support_user.get('id'), build_pending_assignment_message(pending), reply_markup=keyboard)


# ---------------------------------------------------------------------------
# Update handlers
# ---------------------------------------------------------------------------


def handle_message(update):
    message = update.get('message') or {}
    chat = message.get('chat') or {}
    user = message.get('from') or {}
    chat_id = chat.get('id')
    text = message.get('text') or ''

    if not chat_id:
        return False

    def _record(_data):
        _store_user_profile(_data, user)
        return True

    mutate_json(USERS_FILE, default_users(), _record)

    if text == '/export' and is_admin(chat_id):
        write_exports_and_send(chat_id)
        return True

    if text == '/pending' and can_assign_vouchers(chat_id):
        send_pending_summary(user or {'id': chat_id})
        return True

    if is_admin(chat_id):
        if text == '/admin' or text == '/packages':
            pkgs = load_packages()
            lines = ['🛠 *Admin Package Management*', '']
            lines.append('Active Packages:')
            for idx, p in enumerate(pkgs, 1):
                lines.append('%d. *%s* (`%s`) - %s ETB' % (idx, p['label'], p['key'], p['priceCents']/100.0))
            lines.append('')
            lines.append('*Commands:*')
            lines.append('• `/addpackage <key> | <label> | <price_etb>`')
            lines.append('• `/editprice <key> <price_etb>`')
            lines.append('• `/deletepackage <key>`')
            lines.append('• `/addvouchers <pool_key> <code1> <code2> ...`')
            send_message(chat_id, '\n'.join(lines), parse_mode='Markdown')
            return True

        if text == '/addpackage' or text == '/addpackage ':
            _save_draft(chat_id, {'admin_step': 'addpkg_name', 'updatedAt': utcnow()})
            send_message(chat_id, '✏️ Please send the package *Name / Label* (e.g. `Freshman Natural Science - Semester 1`):', parse_mode='Markdown')
            return True

        if text.startswith('/addpackage '):
            raw = text.split(' ', 1)[1].strip()
            parts = [p.strip() for p in raw.split('|')]
            if len(parts) < 3:
                send_message(chat_id, 'Usage: `/addpackage <key> | <label> | <price_etb>` or simply send `/addpackage` for guided setup.', parse_mode='Markdown')
                return True
            key, label, price_str = parts[0], parts[1], parts[2]
            try:
                price_cents = int(float(price_str) * 100)
            except ValueError:
                send_message(chat_id, 'Invalid price amount.')
                return True
            pkgs = load_packages()
            pkgs = [p for p in pkgs if p['key'] != key]
            pkgs.append({
                'key': key,
                'label': label,
                'priceCents': price_cents,
                'currency': 'ETB',
                'phrasePool': key,
            })
            save_packages(pkgs)
            send_message(chat_id, '✅ Package *%s* added/updated successfully.' % label, parse_mode='Markdown')
            return True

        if text.startswith('/editprice '):
            parts = text.split(' ')
            if len(parts) < 3:
                send_message(chat_id, 'Usage: `/editprice <key> <price_etb>`', parse_mode='Markdown')
                return True
            key, price_str = parts[1].strip(), parts[2].strip()
            try:
                price_cents = int(float(price_str) * 100)
            except ValueError:
                send_message(chat_id, 'Invalid price amount.')
                return True
            pkgs = load_packages()
            found = False
            for p in pkgs:
                if p['key'] == key:
                    p['priceCents'] = price_cents
                    found = True
                    break
            if found:
                save_packages(pkgs)
                send_message(chat_id, '✅ Price for package *%s* updated to %s ETB.' % (key, price_str), parse_mode='Markdown')
            else:
                send_message(chat_id, 'Package key not found.')
            return True

        if text.startswith('/deletepackage '):
            key = text.split(' ', 1)[1].strip()
            pkgs = load_packages()
            new_pkgs = [p for p in pkgs if p['key'] != key]
            if len(new_pkgs) < len(pkgs):
                save_packages(new_pkgs)
                send_message(chat_id, '✅ Package *%s* deleted.' % key, parse_mode='Markdown')
            else:
                send_message(chat_id, 'Package key not found.')
            return True

        if text.startswith('/addvouchers '):
            parts = text.split(' ')
            if len(parts) < 3:
                send_message(chat_id, 'Usage: `/addvouchers <pool_key> <code1> <code2> ...`', parse_mode='Markdown')
                return True
            pool_key = parts[1].strip()
            codes = [c.strip() for c in parts[2:] if c.strip()]
            def _mutate_vouchers(data):
                pkgs_store = data.setdefault('packages', {})
                pool_store = pkgs_store.setdefault(pool_key, {'phrases': [], 'issued': {}})
                phrases = pool_store.setdefault('phrases', [])
                for code in codes:
                    if code not in phrases:
                        phrases.append(code)
                return len(codes)
            count = mutate_json(VOUCHERS_FILE, default_vouchers(), _mutate_vouchers)
            send_message(chat_id, '✅ Added %d voucher(s) to pool *%s*.' % (count, pool_key), parse_mode='Markdown')
            return True

    if text.startswith('/revokeaccess ') and is_admin(chat_id):
        token = text.split(' ', 1)[1].strip()
        if revoke_access_token(token):
            send_message(chat_id, 'Access token revoked. The app will no longer open the package.')
        else:
            send_message(chat_id, 'Token not found or already revoked.')
        return True

    if text.startswith('/revokeentitlement ') and is_admin(chat_id):
        entitlement_id = text.split(' ', 1)[1].strip()
        if revoke_entitlement(entitlement_id):
            send_message(chat_id, 'Entitlement revoked. Associated access tokens are no longer valid.')
        else:
            send_message(chat_id, 'Entitlement not found.')
        return True

    if text in ('/start', '/buy'):
        draft = _new_draft(chat_id)
        _save_draft(chat_id, draft)
        send_message(chat_id, get_msg(chat_id, 'lang_select'), reply_markup=build_language_keyboard())
        return True

    if text == '/myaccess':
        voucher = get_voucher_for_user(chat_id)
        if voucher:
            status = normalize_voucher_status(voucher)
            package_label = voucher.get('packageLabel') or ''
            if status == 'redeemed':
                send_message(chat_id, get_msg(chat_id, 'myaccess_redeemed', voucher.get('redeemedAt')))
            else:
                txt, kb = build_approved_message(package_label, voucher.get('phrase'), user_id=chat_id, phone_number=voucher.get('phone'))
                send_message(chat_id, txt, parse_mode='Markdown', reply_markup=kb)
        else:
            send_message(chat_id, get_msg(chat_id, 'no_voucher'))
        return True

    if text == '/diag':
        def _collect(data):
            requests = data.get('requests') or {}
            statuses = [req.get('status') for req in requests.values()]
            latest = [req for req in requests.values()][-3:]
            drafts = [
                (uid, d.get('step'), d.get('package'), d.get('method'),
                 bool(d.get('link')), bool(d.get('txid')), d.get('name'))
                for uid, d in (data.get('drafts') or {}).items()
            ]
            return len(requests), statuses, latest, drafts

        total, statuses, latest, drafts = mutate_json(USERS_FILE, default_users(), _collect)
        lines = [
            'Diag',
            'your chat id: %s' % chat_id,
            'ADMIN_CHAT_ID=%s' % (ADMIN_CHAT_ID or '(not set)'),
            'SUPPORT_CHAT_ID=%s' % (SUPPORT_CHAT_ID or '(not set)'),
            'is_admin=%s' % is_admin(chat_id),
            'DATA_DIR=%s' % DATA_DIR,
            'requests_total=%d' % total,
            'request_statuses=%s' % (', '.join(statuses) or '(none)'),
        ]
        for req in latest:
            lines.append('  %s | %s | %s' % (
                req.get('requestId'), req.get('status'),
                req.get('packageLabel') or req.get('package', '')))
        lines.append('drafts:')
        for uid, step, pkg, method, has_link, has_txid, name in drafts:
            lines.append('  %s | step=%s | pkg=%s | method=%s | link=%s | txid=%s | name=%s' % (
                uid, step, pkg, method, has_link, has_txid, name or ''))
        if not drafts:
            lines.append('  (none)')
        send_message(chat_id, '\n'.join(lines))
        return True

    if is_admin(chat_id):
        admin_draft = _get_draft(chat_id) or {}
        admin_step = admin_draft.get('admin_step')
        if admin_step == 'addpkg_name':
            label = text.strip()
            admin_draft['addpkg_label'] = label
            admin_draft['admin_step'] = 'addpkg_key'
            _save_draft(chat_id, admin_draft)
            send_message(chat_id, '🔑 Got it! Now send the package *Tag / Key* (e.g. `freshman_natural_science_y1_sem1`):', parse_mode='Markdown')
            return True
        elif admin_step == 'addpkg_key':
            key = text.strip().lower().replace(' ', '_')
            admin_draft['addpkg_key'] = key
            admin_draft['admin_step'] = 'addpkg_price'
            _save_draft(chat_id, admin_draft)
            send_message(chat_id, '💰 Great! Finally, send the *Price Tag in ETB* (e.g. `300`):', parse_mode='Markdown')
            return True
        elif admin_step == 'addpkg_price':
            price_str = text.strip()
            try:
                price_cents = int(float(price_str) * 100)
            except ValueError:
                send_message(chat_id, '⚠️ Invalid price amount. Please enter a valid number (e.g. `300`):', parse_mode='Markdown')
                return True
            label = admin_draft.get('addpkg_label') or 'New Package'
            key = admin_draft.get('addpkg_key') or 'new_package'
            pkgs = load_packages()
            pkgs = [p for p in pkgs if p['key'] != key]
            pkgs.append({
                'key': key,
                'label': label,
                'priceCents': price_cents,
                'currency': 'ETB',
                'phrasePool': key,
            })
            save_packages(pkgs)
            _save_draft(chat_id, {})
            send_message(chat_id, '✅ Package successfully added!\n\n📌 *Name:* %s\n🏷 *Tag:* `%s`\n💰 *Price:* %s ETB' % (label, key, price_str), parse_mode='Markdown')
            return True

    if not text or text.startswith('/'):
        return False

    draft, message, reply_markup = handle_draft_step(user, text)
    logger.info('Flow user=%s step=%s', user.get('id'), draft.get('step'))
    if message:
        send_message(chat_id, message, reply_markup=reply_markup)
    elif draft and draft.get('step') == 'done':
        draft['step'] = 'submitted'
        _save_draft(chat_id, draft)
        submit_request(chat_id, draft)
    elif draft and draft.get('step') == 'submitted':
        send_message(chat_id, get_msg(chat_id, 'already_submitted'))
    return True


def handle_contact(update):
    message = update.get('message') or {}
    contact = message.get('contact') or {}
    chat_id = (message.get('chat') or {}).get('id')
    user = message.get('from') or {}

    if not chat_id or not contact:
        return False

    number = contact.get('phone_number') or ''
    if not is_plausible_phone(number):
        send_message(chat_id, get_msg(chat_id, 'invalid_phone'), reply_markup=build_phone_share_keyboard(chat_id))
        return False

    draft = _get_draft(chat_id) or _new_draft(chat_id)
    if draft.get('step') != 'phone':
        return False

    draft.update({
        'phone': {'number': number, 'verified': True},
        'step': 'method',
        'updatedAt': utcnow(),
    })
    _save_draft(chat_id, draft)
    send_message(chat_id, get_msg(chat_id, 'payment_method'), reply_markup=build_payment_method_keyboard())
    return True


def handle_callback(update):
    callback = update.get('callback_query') or {}
    data = callback.get('data') or ''
    user = callback.get('from') or {}
    message = callback.get('message') or {}
    chat_id = (message.get('chat') or {}).get('id')
    message_id = message.get('message_id')
    callback_id = callback.get('id')

    if not data or not callback_id:
        return False

    def _record(_data):
        _store_user_profile(_data, user)
        return True

    mutate_json(USERS_FILE, default_users(), _record)

    if data.startswith('lang:'):
        lang_code = data.split(':', 1)[1]
        set_user_lang(user.get('id'), lang_code)
        answer_callback_query(callback_id, 'Language updated')
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
        draft = _new_draft(user.get('id'))
        _save_draft(user.get('id'), draft)
        send_message(chat_id, get_start_message(user.get('id')), reply_markup=build_package_keyboard())
        return True

    if data.startswith('copy_phone:'):
        phone_num = data.split(':', 1)[1]
        answer_callback_query(callback_id, 'Copy Phone Number')
        send_message(chat_id, get_msg(chat_id, 'copy_phone', phone_num), parse_mode='Markdown')
        return True

    if data.startswith('copy_code:'):
        code_str = data.split(':', 1)[1]
        answer_callback_query(callback_id, 'Copy Redeem Code')
        send_message(chat_id, get_msg(chat_id, 'copy_code', code_str), parse_mode='Markdown')
        return True

    if data.startswith('confirm_ref:'):
        ref = data.split(':', 1)[1]
        answer_callback_query(callback_id, 'Purchase reference confirmed')
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
        handle_draft_step(user, 'confirm_ref:' + ref)
        send_message(chat_id, get_msg(user.get('id'), 'share_phone_btn'), reply_markup=build_phone_share_keyboard(user.get('id')))
        return True

    if data == 'cancel_ref':
        answer_callback_query(callback_id, 'Cancelled')
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
        draft = _new_draft(user.get('id'))
        _save_draft(user.get('id'), draft)
        send_message(chat_id, get_msg(user.get('id'), 'cancelled'))
        return True

    if data == 'cat:main':
        answer_callback_query(callback_id, 'Main Menu')
        if message_id:
            edit_message_text(chat_id, message_id, get_start_message(user.get('id')), reply_markup=build_package_keyboard())
        return True

    if data.startswith('freshman:'):
        stream = data.split(':', 1)[1]
        prefix = 'freshman_natural_science_y1' if stream == 'nat' else 'freshman_social_science_y1'
        stream_name = 'Natural Science' if stream == 'nat' else 'Social Science'
        answer_callback_query(callback_id, stream_name)
        if message_id:
            edit_message_text(chat_id, message_id, '🎓 *Freshman %s (Year 1)*\nSelect your duration / plan:' % stream_name, parse_mode='Markdown', reply_markup=build_plan_keyboard(prefix))
        return True

    if data.startswith('univ_yr:'):
        year = data.split(':', 1)[1]
        answer_callback_query(callback_id, 'Year %s' % year)
        if message_id:
            edit_message_text(chat_id, message_id, '🏛 *University Year %s*\nSelect your Department:' % year, parse_mode='Markdown', reply_markup=build_university_dept_keyboard(year))
        return True

    if data.startswith('univ_dept:'):
        parts = data.split(':')
        dept_key = parts[1]
        year = parts[2].replace('y', '')
        prefix = 'university_%s_y%s' % (dept_key, year)
        dept_label = dept_key.replace('_', ' ').title()
        answer_callback_query(callback_id, dept_label)
        if message_id:
            edit_message_text(chat_id, message_id, '🏛 *%s (Year %s)*\nSelect your duration / plan:' % (dept_label, year), parse_mode='Markdown', reply_markup=build_plan_keyboard(prefix))
        return True

    if data.startswith('package:'):
        package_key = data.split(':', 1)[1]
        if package_key == 'freshman':
            answer_callback_query(callback_id, 'Freshman')
            if message_id:
                edit_message_text(chat_id, message_id, '🎓 *Freshman Year 1*\nSelect your stream:', parse_mode='Markdown', reply_markup=build_freshman_stream_keyboard())
            return True
        if package_key == 'university-department':
            answer_callback_query(callback_id, 'University Department')
            if message_id:
                edit_message_text(chat_id, message_id, '🏛 *University Department*\nSelect your Academic Year:', parse_mode='Markdown', reply_markup=build_university_year_keyboard())
            return True
        if package_key == 'coc' or package_key == 'exit-exam':
            answer_callback_query(callback_id, 'COC / Exit Exam')
            if message_id:
                edit_message_text(chat_id, message_id, '📋 *COC Exam Preparation*\nSelect your exam field:', parse_mode='Markdown', reply_markup=build_coc_keyboard())
            return True

        pcfg = get_product_config(package_key)
        if pcfg:
            package = {'key': pcfg['id'], 'label': pcfg['label'], 'priceCents': pcfg['priceCents'], 'currency': pcfg.get('currency', 'ETB')}
        else:
            package = get_package_by_key(package_key)
        if not package:
            answer_callback_query(callback_id, 'Unknown package')
            return True
        draft = _get_draft(user.get('id')) or _new_draft(user.get('id'))
        draft.update({'package': package['key'], 'step': 'phone', 'updatedAt': utcnow()})
        _save_draft(user.get('id'), draft)
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
        send_message(chat_id,
                     get_msg(user.get('id'), 'package_chosen', package['label'], format_money(package['priceCents'], package['currency'])),
                     reply_markup=build_phone_share_keyboard(user.get('id')))
        return True

    if data.startswith('method:'):
        method_key = data.split(':', 1)[1]
        method = get_payment_method(method_key)
        if not method:
            answer_callback_query(callback_id, 'Unknown method')
            return True
        draft = _get_draft(user.get('id')) or _new_draft(user.get('id'))
        if draft.get('step') != 'method':
            return True
        draft.update({'method': method['key'], 'step': 'name', 'updatedAt': utcnow()})
        _save_draft(user.get('id'), draft)
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
        send_message(chat_id, get_msg(user.get('id'), 'ask_name'))
        return True

    if data.startswith('approve:') and is_admin(user.get('id')):
        request_id = data.split(':', 1)[1]
        req = finalize_request(request_id, 'approved', user)
        if not req:
            answer_callback_query(callback_id, 'Request not found')
            return True

        if message_id:
            edit_message_reply_markup(chat_id, message_id)
            send_message(chat_id, '✅ Approved request `%s`' % request_id, parse_mode='Markdown')

        result = assign_voucher(request_id, user)
        if result in ('ok', 'already'):
            req = get_request_by_id(request_id)
            if req:
                user_phone = (req.get('phone') or {}).get('number') or ''
                txt, kb = build_approved_message(req.get('packageLabel'), req.get('voucher', {}).get('phrase'), user_id=req.get('userId'), phone_number=user_phone)
                send_message(req.get('userId'), txt, parse_mode='Markdown', reply_markup=kb)
            answer_callback_query(callback_id, 'Approved and voucher sent')
        elif result == 'empty':
            send_message(req.get('userId'), build_approved_pending_voucher_message(req))
            answer_callback_query(callback_id, 'Approved, but voucher pool is empty')
        else:
            answer_callback_query(callback_id, 'Request not found')
        return True

    if data.startswith('reject:') and is_admin(user.get('id')):
        request_id = data.split(':', 1)[1]
        req = finalize_request(request_id, 'rejected', user)
        if message_id:
            edit_message_reply_markup(chat_id, message_id)
            send_message(chat_id, '❌ Rejected request `%s`' % request_id, parse_mode='Markdown')
        if req:
            send_message(req.get('userId'), build_rejected_message(req))
        answer_callback_query(callback_id, 'Rejected')
        return True

    if data.startswith('assign:') and can_assign_vouchers(user.get('id')):
        request_id = data.split(':', 1)[1]
        result = assign_voucher(request_id, user)
        if result == 'ok':
            req = get_request_by_id(request_id)
            if req:
                user_phone = (req.get('phone') or {}).get('number') or ''
                txt, kb = build_approved_message(req.get('packageLabel'), req.get('voucher', {}).get('phrase'), user_id=req.get('userId'), phone_number=user_phone)
                send_message(req.get('userId'), txt, parse_mode='Markdown', reply_markup=kb)
            answer_callback_query(callback_id, 'Voucher assigned and sent')
            send_pending_summary(user)
        elif result == 'empty':
            answer_callback_query(callback_id, 'Voucher pool is empty')
        elif result == 'already':
            answer_callback_query(callback_id, 'Already assigned')
        else:
            answer_callback_query(callback_id, 'Request not found')
        return True

    return False


def get_request_by_id(request_id):
    data = read_json(USERS_FILE, default_users())
    for req in (data.get('requests') or {}).values():
        if req.get('requestId') == request_id:
            return req
    return None


def write_exports_and_send(chat_id):
    def _collect(data):
        requests = (data.get('requests') or {}).values()
        by_package = {}
        for req in requests:
            by_package.setdefault(req.get('packageKey'), []).append(req)
        return by_package

    by_package = mutate_json(USERS_FILE, default_users(), _collect)
    written = write_exports(by_package)
    for file_path in written:
        send_document(chat_id, file_path)


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

app = Flask(__name__)

# PythonAnywhere WSGI entry point imports `application` (see wsgi.py).
application = app


@app.route('/')
def index():
    return jsonify({
        'service': 'tinat-bot',
        'status': 'ok',
        'time': utcnow(),
        'webhook_path': '/telegram-webhook',
        'redeem_path': '/api/v1/vouchers/redeem',
    })


@app.route('/health')
def health():
    return jsonify({'ok': True, 'status': 'ok', 'time': utcnow()})


@app.route('/telegram-webhook', methods=['POST'])
def telegram_webhook():
    update = request.get_json(silent=True) or {}
    if not update:
        return '', 200

    update_id = update.get('update_id')
    if record_update(update_id):
        return jsonify({'ok': True, 'duplicate': True})

    try:
        if handle_callback(update) or handle_message(update) or handle_contact(update):
            pass
        return jsonify({'ok': True})
    except Exception:
        logger.exception('Unhandled error while processing update %s', update_id)
        return jsonify({'ok': False, 'error': 'internal error'}), 200


REDEEM_ERROR_HTTP = {
    'INVALID_CODE': 400,
    'NOT_ASSIGNED': 400,
    'ALREADY_REDEEMED': 400,
    'VOUCHER_REVOKED': 400,
    'NOT_OWNER': 400,
    'PAYMENT_NOT_APPROVED': 400,
    'UNAUTHORIZED': 401,
    'RATE_LIMITED': 429,
    'BAD_REQUEST': 400,
}


def _redeem_error(message):
    """Map an internal redeem_code error message to a contract error code."""
    mapping = {
        'No matching voucher found': 'INVALID_CODE',
        'Voucher is not assigned to you yet': 'NOT_ASSIGNED',
        'Voucher already redeemed': 'ALREADY_REDEEMED',
    }
    return mapping.get(message, 'BAD_REQUEST')


@app.route('/api/v1/purchases/create', methods=['POST'])
def create_purchase_api():
    """Endpoint for Android app to initiate a purchase and get a Purchase Reference (e.g. TH-7F92K).
    
    Body: { "productId": "freshman_natural_science_y1_full_year", "userId": "...", "phone": "..." }
    Success: { "success": true, "purchaseReference": "TH-7F92K", "purchase": {...} }
    """
    if request.mimetype not in ('application/json', 'text/json'):
        return jsonify({'success': False, 'error': 'BAD_REQUEST'}), 400

    body = request.get_json(silent=True) or {}
    product_id = str(body.get('productId') or '').strip()
    user_id = str(body.get('userId') or '').strip()
    phone = str(body.get('phone') or '').strip()

    if not product_id:
        return jsonify({'success': False, 'error': 'MISSING_PRODUCT_ID'}), 400

    cfg = get_product_config(product_id)
    if not cfg:
        return jsonify({'success': False, 'error': 'INVALID_PRODUCT_ID'}), 400

    try:
        ref = create_purchase_request(product_id, user_id=user_id, user_phone=phone)
        purchase = get_purchase_by_reference(ref)
        return jsonify({
            'success': True,
            'purchaseReference': ref,
            'purchase': purchase
        }), 200
    except Exception as exc:
        logger.exception('Failed to create purchase request')
        return jsonify({'success': False, 'error': 'SERVER_ERROR'}), 500


@app.route('/api/v1/purchases/<ref>', methods=['GET'])
def get_purchase_api(ref):
    """Retrieve purchase summary by Purchase Reference.
    
    Success: { "success": true, "purchase": {...} }
    """
    purchase = get_purchase_by_reference(ref)
    if not purchase:
        return jsonify({'success': False, 'error': 'NOT_FOUND'}), 404
    return jsonify({
        'success': True,
        'purchase': purchase
    }), 200


def _authenticate(body):
    """Resolve the authenticated Telegram user id from a request body.

    Accepts either a Mini App initData (verified against the bot token) or a
    previously-issued authToken (from /api/v1/auth/login). Returns the user id
    string, or None when unauthenticated.
    """
    auth_token = str(body.get('authToken') or '').strip()
    if auth_token:
        return resolve_auth_token(auth_token)

    user = verify_init_data(body.get('initData'))
    if user is None:
        return None
    return str(user.get('id'))


@app.route('/api/v1/auth/login', methods=['POST'])
def auth_login_api():
    """Log in a Telegram user from the native Android SDK payload.

    Body (the fields the telegram-login-android SDK hands back):
        {id, first_name, last_name, username, auth_date, hash}
    The hash is verified against the bot token, then a short-lived authToken
    is returned that the app passes to /redeem or /status.

    Success: {ok, userId, authToken, expiresAt}
    Errors:  {ok: false, error: UNAUTHORIZED|BAD_REQUEST}
    """
    if request.mimetype not in ('application/json', 'text/json'):
        return jsonify({'ok': False, 'error': 'BAD_REQUEST'}), 400

    body = request.get_json(silent=True) or {}
    user = verify_login_widget(body)
    if user is None:
        return jsonify({'ok': False, 'error': 'UNAUTHORIZED'}), 401

    token = issue_auth_token(user['id'])
    return jsonify({
        'ok': True,
        'userId': user['id'],
        'authToken': token,
        'expiresAt': utcnow(),
    }), 200


@app.route('/api/v1/vouchers/redeem', methods=['POST'])
def redeem_voucher_api():
    """Redeem a voucher owned by the Telegram user in the Mini App.

    Body: { "initData": "...", "code": "G5LE37QI45M2", "deviceId": "..." }
    `phrase` and `voucherPhrase` are accepted as aliases of `code`. The
    initData hash is verified against the bot token before the voucher is
    atomically marked REDEEMED, and an entitlement is recorded.

    Success:  {ok, packageKey, packageLabel, redeemedAt, access}
    Errors:   {ok: false, error: INVALID_CODE|ALREADY_REDEEMED|...}
    """
    if request.mimetype not in ('application/json', 'text/json'):
        return jsonify({'ok': False, 'error': 'BAD_REQUEST'}), 400

    body = request.get_json(silent=True) or {}
    phrase = str(body.get('code') or body.get('phrase') or body.get('voucherPhrase') or '').strip()
    device_id = str(body.get('deviceId') or '')

    if not phrase:
        return jsonify({'ok': False, 'error': 'BAD_REQUEST'}), 400

    owner_id = _authenticate(body)
    if owner_id is None:
        return jsonify({'ok': False, 'error': 'UNAUTHORIZED'}), 401

    if not rate_limit_redeem(owner_id):
        return jsonify({'ok': False, 'error': 'RATE_LIMITED'}), 429

    result = redeem_code(owner_id, phrase, device_id)
    if isinstance(result, dict) and 'error' in result:
        error = _redeem_error(result['error'])
        return jsonify({'ok': False, 'error': error}), REDEEM_ERROR_HTTP[error]

    return jsonify({
        'ok': True,
        'packageKey': result.get('packageKey'),
        'packageLabel': result.get('packageLabel'),
        'redeemedAt': result.get('redeemedAt'),
        'access': {
            'packageKey': result.get('packageKey'),
            'enabled': True,
        },
    }), 200


@app.route('/api/v1/vouchers/status', methods=['POST'])
def voucher_status_api():
    """Check whether a voucher is valid for this user, without consuming it.

    Body: { "initData": "...", "code": "G5LE37QI45M2" }
    Success: {ok: true, status: ASSIGNED|REDEEMED|..., packageKey, packageLabel}
    Errors:  {ok: false, error: INVALID_CODE|NOT_OWNER|...}
    """
    if request.mimetype not in ('application/json', 'text/json'):
        return jsonify({'ok': False, 'error': 'BAD_REQUEST'}), 400

    body = request.get_json(silent=True) or {}
    phrase = str(body.get('code') or body.get('phrase') or body.get('voucherPhrase') or '').strip()

    if not phrase:
        return jsonify({'ok': False, 'error': 'BAD_REQUEST'}), 400

    owner_id = _authenticate(body)
    if owner_id is None:
        return jsonify({'ok': False, 'error': 'UNAUTHORIZED'}), 401

    data = read_json(VOUCHERS_FILE, default_vouchers())
    issued = data.get('issued') or {}
    entry = None
    for candidate in issued.values():
        if str(candidate.get('phrase')) == phrase:
            entry = candidate
            break

    if entry is None:
        return jsonify({'ok': False, 'error': 'INVALID_CODE'}), 400

    status = normalize_voucher_status(entry)
    if status == 'redeemed':
        return jsonify({'ok': False, 'error': 'ALREADY_REDEEMED'}), 400
    if str(entry.get('ownerId')) != str(owner_id):
        return jsonify({'ok': False, 'error': 'NOT_OWNER'}), 400
    if status != 'assigned':
        return jsonify({'ok': False, 'error': 'NOT_ASSIGNED'}), 400

    return jsonify({
        'ok': True,
        'status': 'ASSIGNED',
        'packageKey': entry.get('packageKey'),
        'packageLabel': entry.get('packageLabel'),
    }), 200


ANDROID_ACTIVATE_ERRORS = {
    'Invalid phone number': 'INVALID_PHONE',
    'Missing redeem code': 'BAD_REQUEST',
    'No matching voucher found': 'INVALID_CODE',
    'Voucher already redeemed': 'ALREADY_REDEEMED',
    'Voucher was revoked': 'VOUCHER_REVOKED',
    'Voucher is not assigned to you yet': 'NOT_ASSIGNED',
    'Phone does not match the voucher owner': 'PHONE_MISMATCH',
    'Phone number was not verified through Telegram': 'PHONE_NOT_VERIFIED',
    'Payment was not approved': 'PAYMENT_NOT_APPROVED',
    'Package is invalid': 'PACKAGE_INVALID',
}

ANDROID_ERROR_HTTP = {
    'BAD_REQUEST': 400,
    'INVALID_PHONE': 400,
    'INVALID_CODE': 400,
    'ALREADY_REDEEMED': 400,
    'VOUCHER_REVOKED': 400,
    'NOT_ASSIGNED': 400,
    'PHONE_MISMATCH': 400,
    'PHONE_NOT_VERIFIED': 400,
    'PAYMENT_NOT_APPROVED': 400,
    'PACKAGE_INVALID': 400,
    'RATE_LIMITED': 429,
    'UNAUTHORIZED': 401,
    'TOKEN_EXPIRED': 401,
    'TOKEN_REVOKED': 401,
    'ENTITLEMENT_NOT_FOUND': 404,
}


def _android_activate_error(message):
    return ANDROID_ACTIVATE_ERRORS.get(message, 'BAD_REQUEST')


def _client_ip():
    forwarded = request.headers.get('X-Forwarded-For', '')
    if forwarded:
        return forwarded.split(',')[0].strip()
    return request.remote_addr or 'unknown'


@app.route('/api/v1/android/activate', methods=['POST'])
def android_activate_api():
    """Activate a purchased package from the standalone Android app.

    Body: { "phone": "+2519XXXXXXXX", "code": "BIO-82KF-91XP" }
    The backend verifies everything against its own records (voucher state,
    owner phone, Telegram phone verification, payment approval, package) and
    never trusts app-supplied identity. On success the voucher is marked
    REDEEMED atomically, an entitlement is created, and an app access token
    is returned.

    Success:  {success, accessToken, package: {id, name}}
    Errors:   {success: false, error: INVALID_CODE|PHONE_MISMATCH|...}
    """
    if request.mimetype not in ('application/json', 'text/json'):
        return jsonify({'success': False, 'error': 'BAD_REQUEST'}), 400

    body = request.get_json(silent=True) or {}
    phone = normalize_phone(body.get('phone'))
    phrase = str(body.get('code') or '').strip()
    ip = _client_ip()

    if not phone or not phrase:
        return jsonify({'success': False, 'error': 'BAD_REQUEST'}), 400
    if not is_plausible_phone(phone):
        return jsonify({'success': False, 'error': 'INVALID_PHONE'}), 400

    if not rate_limit_redeem('android-phone:' + phone, ANDROID_RATE_LIMIT_MAX, ANDROID_RATE_LIMIT_WINDOW_MS):
        return jsonify({'success': False, 'error': 'RATE_LIMITED'}), 429
    if not rate_limit_redeem('android-code:' + phrase, ANDROID_RATE_LIMIT_MAX, ANDROID_RATE_LIMIT_WINDOW_MS):
        return jsonify({'success': False, 'error': 'RATE_LIMITED'}), 429
    if not rate_limit_redeem('android-ip:' + ip, ANDROID_RATE_LIMIT_MAX, ANDROID_RATE_LIMIT_WINDOW_MS):
        return jsonify({'success': False, 'error': 'RATE_LIMITED'}), 429

    logger.info('Android activation attempt phone=%s code_len=%d ip=%s', mask_phone(phone), len(phrase), ip)

    result = activate_voucher(phone, phrase)
    if isinstance(result, dict) and 'error' in result:
        error = _android_activate_error(result['error'])
        logger.info('Android activation failed: %s', error)
        return jsonify({'success': False, 'error': error}), ANDROID_ERROR_HTTP[error]

    token = issue_access_token(result)
    return jsonify({
        'success': True,
        'accessToken': token,
        'package': {
            'id': result.get('packageKey'),
            'name': result.get('packageLabel'),
        },
    }), 200


@app.route('/api/v1/android/entitlement', methods=['GET'])
def android_entitlement_api():
    """Look up the package authorized by an app access token.

    Authorization: Bearer <accessToken>
    Success: {success, package: {id, name}, entitlement: {...}}
    Errors:  {success: false, error: UNAUTHORIZED|TOKEN_EXPIRED|TOKEN_REVOKED|
              ENTITLEMENT_NOT_FOUND}
    """
    auth_header = request.headers.get('Authorization', '')
    scheme, _, token = auth_header.partition(' ')
    if scheme.lower() != 'bearer' or not token.strip():
        return jsonify({'success': False, 'error': 'UNAUTHORIZED'}), 401

    status, entry = resolve_access_token(token.strip())
    if status == 'unknown':
        return jsonify({'success': False, 'error': 'UNAUTHORIZED'}), 401
    if status == 'expired':
        return jsonify({'success': False, 'error': 'TOKEN_EXPIRED'}), 401
    if status == 'revoked':
        return jsonify({'success': False, 'error': 'TOKEN_REVOKED'}), 401

    data = read_json(ENTITLEMENTS_FILE, default_entitlements())
    entitlement = (data.get('entitlements') or {}).get(str(entry.get('entitlementId')))
    if entitlement is None or entitlement.get('revoked'):
        error = 'ENTITLEMENT_NOT_FOUND' if entitlement is None else 'TOKEN_REVOKED'
        return jsonify({'success': False, 'error': error}), ANDROID_ERROR_HTTP[error]

    return jsonify({
        'success': True,
        'package': {
            'id': entitlement.get('packageKey'),
            'name': entitlement.get('packageLabel'),
        },
        'entitlement': {
            'entitlementId': entitlement.get('entitlementId'),
            'packageKey': entitlement.get('packageKey'),
            'packageLabel': entitlement.get('packageLabel'),
            'activatedAt': entitlement.get('redeemedAt'),
        },
    }), 200


def register_webhook():
    """Register the Telegram webhook for this PythonAnywhere host."""
    if not TELEGRAM_BOT_TOKEN:
        logger.error('Cannot register webhook: missing TELEGRAM_BOT_TOKEN')
        return None

    if 'PA_WEBSITE_URL' in os.environ:
        base_url = os.environ['PA_WEBSITE_URL'].rstrip('/')
    else:
        username = os.environ.get('PA_USERNAME', '').strip()
        base_url = ('https://%s.pythonanywhere.com' % username) if username else 'http://127.0.0.1:5000'

    webhook_url = base_url + '/telegram-webhook'
    allowed = json.dumps(['message', 'callback_query', 'channel_post'])
    payload = {
        'url': webhook_url,
        'allowed_updates': allowed,
        'drop_pending_updates': False,
    }
    result = tg_call('setWebhook', data=payload)
    if result and result.get('ok'):
        logger.info('Webhook registered at %s', webhook_url)
    else:
        logger.error('Failed to register webhook: %s', result)
    return result


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=int(os.environ.get('PORT', '5000')))

