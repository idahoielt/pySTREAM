#!/bin/bash
# Installs pySTREAM on macOS: creates a Python environment in the .venv folder
# and installs the packages pySTREAM needs. Double-click this file to run it.
cd "$(dirname "$0")"
echo "Installing pySTREAM. This can take a few minutes..."
echo

fail() {
    echo
    echo "$1"
    read -r -p "Press Enter to close this window."
    exit 1
}

command -v python3 >/dev/null 2>&1 ||
    fail "Python was not found. Install Python 3.12 or newer from https://www.python.org/downloads/ and then double-click this file again."

python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)' ||
    fail "pySTREAM needs Python 3.12 or newer, but this is $(python3 --version). Install a newer Python from https://www.python.org/downloads/ and then double-click this file again."

python3 -c 'import tkinter' 2>/dev/null ||
    fail "This Python has no tkinter, which pySTREAM uses for its window. Install Python from https://www.python.org/downloads/ (it includes tkinter) and then double-click this file again."

python3 -m venv .venv || fail "Could not create the Python environment; see the messages above."
.venv/bin/python -m pip install --upgrade pip || fail "The installation failed; check your internet connection and try again."
.venv/bin/python -m pip install -r requirements.txt || fail "The installation failed; check your internet connection and try again."

echo
echo "pySTREAM is installed. Start it by double-clicking start_mac.command"
read -r -p "Press Enter to close this window."
