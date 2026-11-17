#!/usr/bin/env bash
# Install or update outlook-mcp as a hardened systemd service.
# Safe to re-run: keeps /etc/outlook-mcp/outlook-mcp.env, sealed credentials
# and the stored Microsoft sign-in.
set -euo pipefail
[[ $EUID -eq 0 ]] || { echo "run with sudo" >&2; exit 1; }

SRC="$(cd "$(dirname "$0")/.." && pwd)"
APP=/opt/outlook-mcp

python3 -c 'import venv, ensurepip' 2>/dev/null \
  || { echo "Missing venv support: sudo apt install python3-venv" >&2; exit 1; }

echo "==> Service user"
if ! id outlook-mcp >/dev/null 2>&1; then
  useradd --system --no-create-home --home-dir /nonexistent --shell /usr/sbin/nologin outlook-mcp
  echo "    Created system user outlook-mcp"
fi

echo "==> Copying application to $APP"
install -d -m 0755 "$APP" "$APP/app" "$APP/bin"
install -m 0644 "$SRC"/server.py "$SRC"/auth.py "$SRC"/vault.py "$SRC"/graph_client.py \
  "$SRC"/formatting.py "$SRC"/compose.py "$SRC"/extract.py "$SRC"/events.py "$SRC"/authorize.py \
  "$SRC"/requirements.txt "$APP/app/"
install -m 0755 "$SRC"/deploy/authorize.sh "$SRC"/deploy/seal-credential.sh "$APP/bin/"
ln -sf "$APP/bin/authorize.sh" /usr/local/sbin/outlook-mcp-authorize

echo "==> Python environment"
[[ -x "$APP/venv/bin/python" ]] || python3 -m venv "$APP/venv"
"$APP/venv/bin/pip" install --quiet --disable-pip-version-check -r "$APP/app/requirements.txt"
"$APP/venv/bin/python" -m compileall -q "$APP/app"

echo "==> Configuration"
install -d -m 0755 /etc/outlook-mcp
if [[ ! -f /etc/outlook-mcp/outlook-mcp.env ]]; then
  install -m 0644 "$SRC/deploy/outlook-mcp.env.example" /etc/outlook-mcp/outlook-mcp.env
  echo "    Created /etc/outlook-mcp/outlook-mcp.env: set MS_TENANT_ID and MS_CLIENT_ID"
fi
install -d -o outlook-mcp -g outlook-mcp -m 0700 /var/lib/outlook-mcp

echo "==> Credentials"
if [[ ! -f /etc/credstore.encrypted/outlook-mcp.vault_key.cred ]]; then
  "$APP/bin/seal-credential.sh" vault_key --generate
fi

echo "==> systemd unit"
install -m 0644 "$SRC/deploy/outlook-mcp.service" /etc/systemd/system/outlook-mcp.service
systemctl daemon-reload

missing=0
for c in "client_secret" "mcp_token --generate"; do
  if [[ ! -f "/etc/credstore.encrypted/outlook-mcp.${c%% *}.cred" ]]; then
    echo "    Missing credential: sudo $APP/bin/seal-credential.sh $c"
    missing=1
  fi
done
if [[ $missing -eq 0 ]] && systemctl is-active --quiet outlook-mcp; then
  systemctl restart outlook-mcp
  echo "==> Restarted outlook-mcp"
fi
[[ -f /var/lib/outlook-mcp/vault.json.enc ]] \
  || echo "    Not signed in to Microsoft yet: sudo outlook-mcp-authorize"
echo "Done."
