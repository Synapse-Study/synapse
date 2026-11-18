#!/usr/bin/env bash
# Double-clickable macOS launcher for Synapse
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

if [ -f "$DIR/Synapse.app/Contents/MacOS/Synapse" ]; then
    exec "$DIR/Synapse.app/Contents/MacOS/Synapse" "$@"
elif [ -f "$DIR/.venv/bin/python3" ]; then
    exec "$DIR/.venv/bin/python3" "$DIR/anki_study_suite.py" "$@"
else
    exec python3 "$DIR/anki_study_suite.py" "$@"
fi
