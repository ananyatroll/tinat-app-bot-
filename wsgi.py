# PythonAnywhere WSGI configuration.
# On the PythonAnywhere Web tab, set:
#   Source code and working directory: /home/<username>/tinat-bot
#   WSGI configuration file: /home/<username>/tinat-bot/wsgi.py

import os
import sys

PROJECT_HOME = os.path.expanduser('~/tinat-bot')
if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

from flask_app import application, logger, register_webhook  # noqa: E402, F401

# Keep the Telegram webhook pointed at this host. The webhook is stored on
# Telegram's side, but re-registering on every reload makes the setup
# self-healing: if the bot token or the hostname ever changes, or someone
# re-points the webhook during local testing, the next reload fixes it.
# Only runs on PythonAnywhere (PA_USERNAME / PA_WEBSITE_URL are set there).
if os.environ.get('PA_USERNAME') or os.environ.get('PA_WEBSITE_URL'):
    try:
        register_webhook()
    except Exception:
        logger.exception('Failed to register Telegram webhook at startup')
