#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -x ".venv/bin/python" ]; then
  echo "[1/2] Criando ambiente virtual..."
  python3 -m venv .venv
  echo "[2/2] Instalando dependencias..."
  .venv/bin/python -m pip install --upgrade pip
  .venv/bin/python -m pip install -r requirements.txt
fi
echo "Iniciando RH Valoracao..."
echo "Abra no navegador: http://127.0.0.1:8000"
.venv/bin/python -m uvicorn backend.app:app --host 127.0.0.1 --port 8000
