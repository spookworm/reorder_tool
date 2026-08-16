#!/usr/bin/env bash

set -e

cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
    echo "Creating Python virtual environment..."
    python3 -m venv .venv

    echo "Installing dependencies..."
    .venv/bin/python -m pip install --upgrade pip
    .venv/bin/python -m pip install -r requirements.txt
fi

echo "Starting Goodreads Ranker..."

.venv/bin/python -m streamlit run app.py