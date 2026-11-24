#!/usr/bin/env bash
# Sobe o app no Linux. Por padrão escuta só em 127.0.0.1 (acesse por túnel SSH
# ou por um proxy reverso com senha). Para mudar: HOST=0.0.0.0 PORT=8501 ./run.sh
set -euo pipefail
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install --upgrade pip
  .venv/bin/pip install -r requirements.txt
fi

exec .venv/bin/streamlit run app.py \
  --server.address "${HOST:-127.0.0.1}" \
  --server.port "${PORT:-8501}" \
  --server.headless true \
  --browser.gatherUsageStats false
