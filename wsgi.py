# PythonAnywhere WSGI configuration.
# On the PythonAnywhere Web tab, set:
#   Source code and working directory: /home/<username>/tinat-bot
#   WSGI configuration file: /home/<username>/tinat-bot/wsgi.py

import os
import sys

PROJECT_HOME = os.path.expanduser('~/tinat-bot')
if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

from flask_app import application  # noqa: E402, F401
