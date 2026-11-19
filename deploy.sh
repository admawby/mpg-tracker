#!/bin/sh
# Copy the app to a Docker host and rebuild it there. Leaves data/ alone.
# Needs DEPLOY_HOST (user@host) and DEPLOY_DIR, from the environment or a
# gitignored .deploy.env next to this script.
set -e
cd "$(dirname "$0")"
[ -f .deploy.env ] && . ./.deploy.env
: "${DEPLOY_HOST:?set DEPLOY_HOST=user@host}" "${DEPLOY_DIR:?set DEPLOY_DIR=path/on/host}"
ssh "$DEPLOY_HOST" "mkdir -p '$DEPLOY_DIR'"
scp -q -r Dockerfile compose.yml README.md app.py static "$DEPLOY_HOST:$DEPLOY_DIR/"
ssh "$DEPLOY_HOST" "cd '$DEPLOY_DIR' && docker compose up -d --build 2>&1 | tail -1"
