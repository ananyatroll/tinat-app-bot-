"""Local runner for the Tinat bot.

PythonAnywhere hosted the bot in Telegram webhook mode, which needs a public
URL. On a laptop there is no public URL, so this runs the bot in long-polling
mode (getUpdates) and serves the redeem API on http://127.0.0.1:5000 in a
background thread so the Mini App can still redeem vouchers locally.

Usage:
    python run_local.py

Stops with Ctrl+C.
"""

import threading
import time

import flask_app

POLL_TIMEOUT_SECONDS = 50
ALLOWED_UPDATES = ('message', 'callback_query')
RETRY_DELAY_SECONDS = 5


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
    host = flask_app.os.environ.get('LOCAL_HOST', '127.0.0.1')
    port = int(flask_app.os.environ.get('PORT', '5000'))

    server_thread = threading.Thread(
        target=flask_app.app.run,
        kwargs={'host': host, 'port': port, 'use_reloader': False},
        daemon=True,
        name='flask-redeem-api',
    )
    server_thread.start()
    flask_app.logger.info('Local server (redeem API / health): http://%s:%s', host, port)

    try:
        poll_forever()
    except KeyboardInterrupt:
        flask_app.logger.info('Stopped by user.')


if __name__ == '__main__':
    main()
