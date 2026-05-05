# gunicorn.conf.py
# Optimised for Render (free / starter tier) hosting a Flask Excel-processing app.
# Render sets the PORT environment variable automatically.

import os

# ── Binding ────────────────────────────────────────────────────────────────────
bind = f"0.0.0.0:{os.environ.get('PORT', '5050')}"

# ── Workers ────────────────────────────────────────────────────────────────────
# Render free tier has 512 MB RAM.  Keep workers low to avoid OOM when large
# Excel files are held in SESSION_STORE (in-memory).
workers = 2                 # increase to 4 on a paid instance
worker_class = "sync"       # sync is safest for in-memory SESSION_STORE
threads = 2                 # lightweight concurrency per worker

# ── Timeouts ──────────────────────────────────────────────────────────────────
# Excel processing can be slow for large files; give it up to 120 s.
timeout = 120
keepalive = 5

# ── Upload size ────────────────────────────────────────────────────────────────
# Must be at least as large as Flask's MAX_CONTENT_LENGTH (200 MB).
limit_request_line = 0
limit_request_fields = 200
limit_request_field_size = 0

# ── Logging ───────────────────────────────────────────────────────────────────
accesslog = "-"     # stdout  → visible in Render's log dashboard
errorlog  = "-"     # stderr
loglevel  = "info"

# ── Process naming ────────────────────────────────────────────────────────────
proc_name = "sjc_report_processor"
