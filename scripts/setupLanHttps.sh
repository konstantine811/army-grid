#!/usr/bin/env bash
# Install a local CA (mkcert) and issue a trusted cert for localhost + LAN IPs.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CERT_DIR="$ROOT/.cert"
mkdir -p "$CERT_DIR"

export PATH="$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"

if ! command -v mkcert >/dev/null 2>&1; then
  echo "Немає mkcert. Встановіть: brew install mkcert nss" >&2
  echo "або покладіть бінарник у ~/.local/bin/mkcert" >&2
  exit 1
fi

echo "Додаю локальний CA в системне сховище (Keychain)…"
if ! mkcert -install; then
  echo
  echo "Потрібен пароль адміністратора. Відкрийте звичайний Термінал і виконайте:"
  echo "  export PATH=\"\$HOME/.local/bin:\$PATH\""
  echo "  mkcert -install"
  echo
fi

IPS="$(ifconfig -l | xargs -n1 ipconfig getifaddr 2>/dev/null || true)"
HOSTS=(localhost 127.0.0.1 ::1)
while IFS= read -r ip; do
  [[ -n "$ip" ]] && HOSTS+=("$ip")
done <<< "$IPS"

echo "Випускаю сертифікат: ${HOSTS[*]}"
mkcert \
  -key-file "$CERT_DIR/key.pem" \
  -cert-file "$CERT_DIR/cert.pem" \
  "${HOSTS[@]}"

printf '%s\n' "${HOSTS[@]}" | tr ' ' '\n' | awk 'NF' | sort -u | paste -sd, - > "$CERT_DIR/san.txt"
cp "$(mkcert -CAROOT)/rootCA.pem" "$CERT_DIR/rootCA.pem"

echo
echo "Готово. Перезапустіть Vite і повністю закрийте Chrome (⌘Q)."
echo "Цей Mac: https://localhost:5173 — має бути Secure."
echo "Інші в мережі: встановіть один раз $CERT_DIR/rootCA.pem як довірений корінь."
echo "  iPhone: AirDrop/Files → профіль → Settings → General → About → Certificate Trust Settings"
echo "  Android: Settings → Security → Encryption & credentials → Install a certificate → CA"
echo "  Windows: подвійний клік rootCA.pem → Local Machine → Trusted Root Certification Authorities"
