#!/usr/bin/env bash
# Instala (ou atualiza) o app como serviço systemd no Debian/Ubuntu.
#
#   ./install.sh                  # escuta em 127.0.0.1:8501
#   PORTA=8600 ./install.sh       # outra porta
#   HOST=0.0.0.0 ./install.sh     # aceita conexões de fora (cuidado: o app não tem login)
#
# Rode como seu usuário normal (o script pede sudo quando precisa). O serviço roda com esse
# usuário. Para atualizar depois de um "git pull", basta rodar o script de novo.
set -euo pipefail

SERVICO=status-report
HOST="${HOST:-127.0.0.1}"
PORTA="${PORTA:-${PORT:-8501}}"
PASTA="$(cd "$(dirname "$0")" && pwd)"
USUARIO="$(id -un)"

if [ "$(id -u)" -eq 0 ]; then
  echo "Não rode como root: rode como o usuário que vai executar o app (o script usa sudo)." >&2
  exit 1
fi
case "$PASTA" in
  *" "*) echo "A pasta do app não pode ter espaços: $PASTA" >&2; exit 1 ;;
esac
command -v apt-get >/dev/null || { echo "Este script é para Debian/Ubuntu (apt)." >&2; exit 1; }
command -v systemctl >/dev/null || { echo "systemd não encontrado." >&2; exit 1; }

echo "==> Pacotes do sistema (Python, LibreOffice para a prévia, fontes)"
sudo apt-get update -qq
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y -qq --no-install-recommends \
  python3 python3-venv python3-pip curl \
  libreoffice-impress fonts-crosextra-carlito fonts-liberation fonts-dejavu-core

python3 - <<'PY' || { echo "É preciso Python 3.11 a 3.14." >&2; exit 1; }
import sys
sys.exit(0 if (3, 11) <= sys.version_info[:2] <= (3, 14) else 1)
PY

echo "==> Ambiente Python (.venv): baixa as dependências, pode levar alguns minutos"
[ -d "$PASTA/.venv" ] || python3 -m venv "$PASTA/.venv"
"$PASTA/.venv/bin/pip" install -q --upgrade pip
"$PASTA/.venv/bin/pip" install --progress-bar on -r "$PASTA/requirements.txt"

echo "==> Serviço systemd ($SERVICO)"
sed -e "s|@USUARIO@|$USUARIO|g" -e "s|@PASTA@|$PASTA|g" \
    -e "s|@HOST@|$HOST|g" -e "s|@PORTA@|$PORTA|g" \
    "$PASTA/deploy/status-report.service" | sudo tee "/etc/systemd/system/$SERVICO.service" >/dev/null
sudo systemctl daemon-reload
sudo systemctl enable -q "$SERVICO"
sudo systemctl restart "$SERVICO"

echo "==> Conferindo se o app subiu"
ALVO="$HOST"; [ "$HOST" = "0.0.0.0" ] && ALVO=127.0.0.1
for _ in $(seq 1 30); do
  if curl -fs "http://$ALVO:$PORTA/_stcore/health" >/dev/null; then
    echo "Pronto: app rodando em http://$HOST:$PORTA (serviço '$SERVICO')."
    echo "Logs: journalctl -u $SERVICO -f"
    exit 0
  fi
  sleep 1
done
echo "O app não respondeu em 30 s. Veja: journalctl -u $SERVICO -n 50" >&2
exit 1
