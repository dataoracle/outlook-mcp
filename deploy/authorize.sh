#!/usr/bin/env bash
# Sign in to Microsoft once (delegated OAuth2: authorization code + PKCE) and store
# the refresh token for outlook-mcp, encrypted with the TPM-sealed vault key.
# Run it again whenever the server reports that the sign-in is no longer valid.
#
#   sudo outlook-mcp-authorize          (installed by install.sh)
#   sudo ./deploy/authorize.sh          (from a clone of the repo)
set -euo pipefail
[[ $EUID -eq 0 ]] || { echo "run with sudo" >&2; exit 1; }

APP=/opt/outlook-mcp
ENV_FILE=/etc/outlook-mcp/outlook-mcp.env
KEY=/etc/credstore.encrypted/outlook-mcp.vault_key.cred
SECRET=/etc/credstore.encrypted/outlook-mcp.client_secret.cred
STATE=/var/lib/outlook-mcp
HERE="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"

[[ -x "$APP/venv/bin/python" ]] || { echo "Not installed yet: sudo ./deploy/install.sh" >&2; exit 1; }
[[ -f "$ENV_FILE" ]] || { echo "Missing $ENV_FILE: run install.sh" >&2; exit 1; }
id outlook-mcp >/dev/null 2>&1 || { echo "Missing user outlook-mcp: run install.sh" >&2; exit 1; }
[[ -f "$SECRET" ]] || { echo "Seal the app's client secret first: sudo $HERE/seal-credential.sh client_secret" >&2; exit 1; }
if [[ ! -f "$KEY" ]]; then
  echo "==> No vault key yet, generating one"
  "$HERE/seal-credential.sh" vault_key --generate
fi
install -d -o outlook-mcp -g outlook-mcp -m 0700 "$STATE"

cd "$APP/app"
"$APP/venv/bin/python" "$APP/app/authorize.py" \
  --env-file "$ENV_FILE" --key-cred "$KEY" --secret-cred "$SECRET" \
  --vault "$STATE/vault.json.enc" --owner outlook-mcp

if ! systemctl is-active --quiet outlook-mcp; then
  echo "outlook-mcp is not running. Start it with: sudo systemctl enable --now outlook-mcp"
fi
