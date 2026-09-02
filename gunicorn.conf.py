"""Gunicorn production config. Start with:  gunicorn app:app -c gunicorn.conf.py"""
import os

bind = f"0.0.0.0:{os.environ.get('PORT', '8080')}"

# IMPORTANT: keep workers = 1.
# The PKG Counter background-job store (JOBS in app.py) lives in this process's
# memory. A second worker cannot see those jobs, so status polling would 404
# at random. Use threads for concurrency instead of workers.
# See documents/UPGRADE_PLAN.md section 6 for the fix (JobStore interface).
workers = 1
threads = int(os.environ.get("GUNICORN_THREADS", "4"))
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "180"))

accesslog = "-"
errorlog = "-"
