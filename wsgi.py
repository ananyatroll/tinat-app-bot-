# PythonAnywhere WSGI configuration (kept for reference only).
# The bot now runs locally in long-polling mode via run_local.py, so the
# webhook is intentionally NOT registered here anymore. Re-registering it
# would steal updates away from the local poller (Telegram returns a 409
# conflict while a webhook is set).

import os
import sys

PROJECT_HOME = os.path.expanduser('~/tinat-bot')
if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

from flask_app import application  # noqa: E402, F401
