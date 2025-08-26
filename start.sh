#!/bin/zsh

set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"

echo "[0/6] Ensuring mkcert is installed and generating fresh TLS certs..."
mkdir -p "$PROJECT_ROOT/certs"
if ! command -v mkcert >/dev/null 2>&1; then
  echo "Installing mkcert (requires Homebrew)..."
  if ! command -v brew >/dev/null 2>&1; then
    echo "Homebrew is required to auto-install mkcert. Install Homebrew from https://brew.sh and re-run." >&2
    exit 1
  fi
  brew install mkcert >/dev/null
fi

# Install local root CA (no-op if already installed)
mkcert -install >/dev/null

# Always create fresh certs per run
rm -f "$PROJECT_ROOT/certs/localhost.crt" "$PROJECT_ROOT/certs/localhost.key"
mkcert -key-file "$PROJECT_ROOT/certs/localhost.key" -cert-file "$PROJECT_ROOT/certs/localhost.crt" localhost 127.0.0.1 ::1 >/dev/null

echo "[1/6] Building and starting docker-compose stack..."
docker compose -f "$PROJECT_ROOT/docker-compose.yaml" up -d --build

echo "[2/6] Waiting for Ollama (via nginx https://localhost/ollama) to be ready..."
until [ "$(curl -sk -o /dev/null -w "%{http_code}" https://localhost/ollama/api/tags)" = "200" ]; do
  sleep 2
  echo "... still waiting"
done

MODEL_NAME=${1:-"gemma3:1b"}
echo "[3/6] Pulling Ollama model: $MODEL_NAME"
OLLAMA_CID=$(docker compose -f "$PROJECT_ROOT/docker-compose.yaml" ps -q ollama)
if [ -z "$OLLAMA_CID" ]; then
  echo "Failed to resolve Ollama container ID" >&2
  exit 1
fi
docker exec "$OLLAMA_CID" ollama pull "$MODEL_NAME"

echo "[4/6] Validating Office Add-in manifest..."
(cd "$PROJECT_ROOT/app" && npx --yes office-addin-manifest validate manifest.xml | cat)

echo "[5/6] Reloading nginx to pick up new certs..."
docker exec ollama_support-nginx-1 nginx -s reload 2>/dev/null || true

echo "[6/6] Sideloading Word add-in (macOS)..."
(cd "$PROJECT_ROOT/app" && npx --yes office-addin-debugging start manifest.xml --app word --platform desktop --no-debugging | cat)

echo "Done. Navigate to https://localhost to load the add-in."

