#!/usr/bin/env bash
# Startup script para Azure App Service (Linux Python).
# Azure escucha en $PORT (típicamente 8000). Streamlit se vincula a 0.0.0.0.
set -euo pipefail

python -m streamlit run app.py \
  --server.port "${PORT:-8000}" \
  --server.address 0.0.0.0 \
  --server.headless true \
  --server.enableCORS false \
  --server.enableXsrfProtection false \
  --browser.gatherUsageStats false
