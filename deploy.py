"""
One-shot deployment helper for PythonAnywhere.

Run from the PythonAnywhere Bash console after uploading the code:

    python ~/tinat-bot/deploy.py

What it does:
  1. Imports flask_app (proves the WSGI module loads cleanly).
  2. Registers the Telegram webhook for this host.
  3. Prints getWebhookInfo so you can confirm the webhook is live.

Required env (Web tab -> Environment variables, or ~/tinat-bot/.env):
  TELEGRAM_BOT_TOKEN or BOT_TOKEN
  PA_USERNAME (or PA_WEBSITE_URL)

On every web-app reload the webhook is re-registered automatically by
wsgi.py, so this script is only needed for the first setup.
"""

import os
import sys

PROJECT_HOME = os.path.expanduser('~/tinat-bot')
sys.path.insert(0, PROJECT_HOME)

from flask_app import register_webhook, tg_call  # noqa: E402


def main():
    if not (os.environ.get('PA_USERNAME') or os.environ.get('PA_WEBSITE_URL')):
        print('Set PA_USERNAME (or PA_WEBSITE_URL) first: '
              'Web tab -> Environment variables.')
        sys.exit(1)

    result = register_webhook()
    if not result or not result.get('ok'):
        print('Webhook registration FAILED:', result)
        sys.exit(1)
    print('Webhook registered:', result)

    info = tg_call('getWebhookInfo')
    if info:
        print('getWebhookInfo:', info)
        if not info.get('ok'):
            sys.exit(1)
        print('OK - bot is live at', info['result'].get('url'))

    print('Deployment OK. Check https://<username>.pythonanywhere.com/health too.')


if __name__ == '__main__':
    main()
