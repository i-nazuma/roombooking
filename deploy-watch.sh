#!/usr/bin/env bash
# Runs server.py and checks origin/main every few minutes for new commits.
# On a new commit: force-syncs to origin/main (git reset --hard — safe
# here, since config.json/token caches are gitignored and untouched by it)
# and restarts the server. Meant for the tablet; keep this running instead
# of `python server.py` directly.
#
# Usage (from the repo root):
#   nohup ./deploy-watch.sh > deploy-watch.log 2>&1 &

set -u
CHECK_INTERVAL_SECONDS=300  # 5 minutes — a git fetch is cheap enough to check this often

cd "$(dirname "$0")"

start_server() {
  python server.py &
  SERVER_PID=$!
  echo "$(date '+%Y-%m-%d %H:%M:%S') server started (pid $SERVER_PID)"
}

restart_server() {
  if [ -n "${SERVER_PID:-}" ] && kill -0 "$SERVER_PID" 2>/dev/null; then
    kill "$SERVER_PID"
    wait "$SERVER_PID" 2>/dev/null
  fi
  start_server
}

start_server

while true; do
  sleep "$CHECK_INTERVAL_SECONDS"

  git fetch origin main --quiet
  LOCAL=$(git rev-parse HEAD)
  REMOTE=$(git rev-parse origin/main)

  if [ "$LOCAL" != "$REMOTE" ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') new commits on origin/main ($LOCAL -> $REMOTE), updating"
    git reset --hard origin/main --quiet
    pip install -q -r requirements.txt
    restart_server
  fi
done
