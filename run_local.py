"""Runner for the Tinat bot (local laptop, Railway, or any host).

This runs the bot in long-polling mode (getUpdates), which needs no public
URL and works on a laptop as well as on Railway, and serves the redeem API on
0.0.0.0:$PORT in a background thread so the Mini App can redeem vouchers.

Environment:
  HOST (default 0.0.0.0)     interface to bind the Flask server to
  PORT (default 5000)        port to listen on (Railway sets this)
  DATA_DIR                   where state JSON lives; on Railway point this
                             at a mounted persistent volume, e.g. /data

Usage:
    python run_local.py

Stops with Ctrl+C.
"""

import os
import threading
import time

import flask_app

POLL_TIMEOUT_SECONDS = 50
ALLOWED_UPDATES = ('message', 'callback_query')
RETRY_DELAY_SECONDS = 5


def _lock_path():
    return os.path.join(flask_app.DATA_DIR, 'run_local.lock')


def acquire_poll_lock():
    """Ensure only one getUpdates poller runs for this bot.

    Telegram itself rejects a second poller with error 409, so this lock only
    needs to keep a single instance alive on the shared volume. A leftover
    lock from a previous container is always taken over immediately: on
    Railway every container runs as pid 1, so the recorded owner pid is never
    meaningful across restarts, and waiting for a 'stale' window just turns a
    restart into a crash loop.
    """
    path = _lock_path()
    try:
        with open(path, 'x', encoding='utf-8') as fh:
            fh.write(str(os.getpid()))
        flask_app.logger.info('Acquired single-instance poll lock: %s', path)
        return path
    except FileExistsError:
        pass

    try:
        owner = open(path, encoding='utf-8').read().strip()
    except OSError:
        owner = 'unknown'
    age = time.time() - os.path.getmtime(path)
    flask_app.logger.warning(
        'Leftover poll lock found (pid %s, %ds old) from a previous run; '
        'taking it over. If another instance of this bot is still polling, '
        'Telegram will reject its polls with 409.', owner, int(age))
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(str(os.getpid()))
    return path


def release_poll_lock(path):
    try:
        with open(path, encoding='utf-8') as fh:
            owner = fh.read().strip()
    except OSError:
        return
    if owner == str(os.getpid()):
        try:
            os.remove(path)
        except OSError:
            pass


def touch_poll_lock():
    try:
        os.utime(_lock_path())
    except OSError:
        pass


def clear_webhook():
    flask_app.logger.info('Clearing any Telegram webhook before polling ...')
    result = flask_app.tg_call('deleteWebhook', data={'drop_pending_updates': True})
    if result and result.get('ok'):
        flask_app.logger.info('Webhook cleared (drop_pending_updates=True).')
    else:
        flask_app.logger.warning('deleteWebhook did not report ok; continuing anyway.')


def verify_bot_token():
    """Confirm the bot token is set and identify which bot is being polled.

    Returns True when the token works, otherwise logs a clear error. Without a
    working BOT_TOKEN the poller silently does nothing, so this makes the
    failure obvious in the logs.
    """
    result = flask_app.tg_call('getMe')
    if not result or not result.get('ok'):
        description = (result or {}).get('description') or ''
        flask_app.logger.error(
            'Telegram getMe failed. BOT_TOKEN is missing or invalid '
            '(set BOT_TOKEN on Railway/Variables and redeploy): %s', description)
        return False
    info = result.get('result') or {}
    flask_app.logger.info(
        'Connected to Telegram as @%s (id %s). Polling will deliver updates '
        'to this bot.', info.get('username'), info.get('id'))
    return True


def log_storage_paths():
    flask_app.logger.info('DATA_DIR: %s', flask_app.DATA_DIR)
    for label in ('USERS_FILE', 'VOUCHERS_FILE', 'ENTITLEMENTS_FILE',
                  'PROCESSED_UPDATES_FILE', 'AUTH_TOKENS_FILE', 'ACCESS_TOKENS_FILE'):
        flask_app.logger.info('Storage %s: %s', label, getattr(flask_app, label))


def get_updates(offset, timeout=POLL_TIMEOUT_SECONDS):
    return flask_app.tg_call('getUpdates', timeout=timeout + 10, json={
        'offset': offset,
        'timeout': timeout,
        'allowed_updates': ALLOWED_UPDATES,
    })


def process_update(update):
    update_id = update.get('update_id')
    if flask_app.record_update(update_id):
        flask_app.logger.info('Skipped duplicate update %s', update_id)
        return

    try:
        if (flask_app.handle_callback(update)
                or flask_app.handle_message(update)
                or flask_app.handle_contact(update)):
            flask_app.logger.info('Processed update %s', update_id)
        else:
            flask_app.logger.info('Unhandled update %s', update_id)
    except Exception:
        flask_app.logger.exception('Error processing update %s', update_id)


def poll_forever():
    offset = 0
    flask_app.logger.info('Starting long polling (getUpdates mode) ...')
    while True:
        touch_poll_lock()
        try:
            result = get_updates(offset)
        except Exception:
            flask_app.logger.exception('getUpdates failed')
            time.sleep(RETRY_DELAY_SECONDS)
            continue

        if not result or not result.get('ok'):
            error_code = (result or {}).get('error_code')
            description = (result or {}).get('description')
            flask_app.logger.warning('getUpdates error %s: %s', error_code, description)
            if error_code == 409:
                flask_app.logger.warning(
                    'A webhook is still set for this bot. Remove it with:\n'
                    '  python -c "import flask_app as f; '
                    'f.tg_call(\'deleteWebhook\', data={\'drop_pending_updates\': True})"')
            time.sleep(RETRY_DELAY_SECONDS)
            continue

        updates = result.get('result') or []
        for update in updates:
            process_update(update)
            update_id = update.get('update_id')
            if update_id is not None:
                offset = max(offset, update_id + 1)


def main():
    lock_path = acquire_poll_lock()
    log_storage_paths()

    host = flask_app.os.environ.get('HOST', '0.0.0.0')
    port = int(flask_app.os.environ.get('PORT', '5000'))

    server_thread = threading.Thread(
        target=flask_app.app.run,
        kwargs={'host': host, 'port': port, 'threaded': True, 'use_reloader': False},
        daemon=True,
        name='flask-redeem-api',
    )
    server_thread.start()
    flask_app.logger.info('Local server (redeem API / health): http://%s:%s', host, port)

    clear_webhook()

    try:
        if not verify_bot_token():
            raise RuntimeError('BOT_TOKEN check failed; refusing to poll.')
        poll_forever()
    except KeyboardInterrupt:
        flask_app.logger.info('Stopped by user.')
    finally:
        release_poll_lock(lock_path)


if __name__ == '__main__':
    main()
