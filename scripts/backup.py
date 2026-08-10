"""
Scheduled backup for the Tinat bot data stores.

Run on PythonAnywhere as a scheduled task (Tasks tab, Bash):
    python ~/tinat-bot/scripts/backup.py

Produces ~/tinat-bot/backups/tinat_<name>_<timestamp>.zip for the
data/ and exports/ folders. Keeps the newest BACKUP_KEEP backups
(default 14) and removes older ones. Safe to run while the bot is live:
JSON stores are written atomically (tmp file + os.replace), so a backup
captures a consistent snapshot.

Optional env: BACKUP_KEEP (int, default 14)
"""

import glob
import os
import shutil
import sys
import time

BASE_DIR = os.path.expanduser('~/tinat-bot')
DATA_DIR = os.path.join(BASE_DIR, 'data')
EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')
BACKUP_DIR = os.path.join(BASE_DIR, 'backups')
BACKUP_KEEP = int(os.environ.get('BACKUP_KEEP', '14'))


def backup_dir(name, source_dir):
    if not os.path.isdir(source_dir):
        print('skip %s: %s not found' % (name, source_dir))
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    base = os.path.join(BACKUP_DIR, 'tinat_%s_%s' % (name, timestamp))
    path = shutil.make_archive(base, 'zip', source_dir)
    print('backed up %s -> %s' % (name, path))
    return path


def prune():
    files = sorted(glob.glob(os.path.join(BACKUP_DIR, 'tinat_*.zip')))
    for old in files[:-BACKUP_KEEP] if len(files) > BACKUP_KEEP else []:
        os.remove(old)
        print('removed old backup %s' % old)


def main():
    backup_dir('data', DATA_DIR)
    backup_dir('exports', EXPORTS_DIR)
    prune()
    print('Backup done.')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:  # keep scheduled-task output visible
        print('Backup FAILED:', exc)
        sys.exit(1)
