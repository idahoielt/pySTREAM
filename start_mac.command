#!/bin/bash
# Starts pySTREAM on macOS. Double-click this file to run it.
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
    echo "pySTREAM is not installed yet. Double-click install_mac.command first."
    read -r -p "Press Enter to close this window."
    exit 1
fi
echo "Starting pySTREAM. The window can take up to a minute to appear..."
if ! .venv/bin/python pySTREAM.py; then
    # Keep this window open if pySTREAM stopped with an error, so the message can be read
    read -r -p "pySTREAM stopped with an error (see above). Press Enter to close this window."
fi
