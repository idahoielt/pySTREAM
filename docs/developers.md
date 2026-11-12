# pySTREAM for developers

This page is for people who want to change pySTREAM or run it without the install
scripts. Students only need the [README](../README.md).

## Folders

```
pySTREAM/
├── program/
│   ├── pySTREAM.py       # the window (tkinter); start pySTREAM with this file
│   ├── pySTREAM_mac.py   # the window, macOS version (used by start_mac.command)
│   ├── Backend.py        # the model: energy balance, CO2 and costs
│   └── Time_Series.py    # the hourly time series
├── data/                 # input data and the current scenario (see the README)
├── docs/                 # this page, the changelog and the comparison with Excel STREAM
├── tests/                # automatic tests
├── requirements.txt      # Python packages
├── environment.yml       # optional conda environment (uses requirements.txt)
├── install_*/start_*     # install and start scripts for Windows and macOS
└── figures/              # plots saved while running (created automatically)
```

The data, the figures and the error log (`pySTREAM error log.txt`) are always in the main
pySTREAM folder, whichever folder pySTREAM is started from.

## Two versions of the window

`program/pySTREAM_mac.py` is a copy of `program/pySTREAM.py` with changes for macOS
(buttons that keep their colours, plots drawn with Tk, text scaled with the screen). It is
developed on the `mac` branch, where it is called `pySTREAM.py`. Changes to the window
must be made in both files. Both use the same model (`Backend.py`, `Time_Series.py`).

The automatic tests on this branch test `program/pySTREAM.py`; the macOS version is
tested on the `mac` branch.

## Setting up without the install scripts

You need Python 3.12 or newer. From the pySTREAM folder:

```bash
python -m venv .venv
.venv\Scripts\activate                # Windows
source .venv/bin/activate             # macOS/Linux
pip install -r requirements.txt
python program/pySTREAM.py            # Windows/Linux
python program/pySTREAM_mac.py        # macOS
```

On Linux, tkinter may need to be installed separately (e.g. `sudo apt install python3-tk`).
pySTREAM has not been tested on Linux.

### With conda

See [Install with conda](../README.md#install-with-conda-optional) in the README.
`environment.yml` takes Python from conda-forge (with tkinter) and installs the packages
from `requirements.txt` with pip, so the package versions are only listed in
`requirements.txt`.

## Tests

The tests check that results do not change unintentionally, that they match the Excel
STREAM model, and that the main parts of the window work. Run them from the pySTREAM
folder with the environment active (takes a few minutes):

```bash
python -m unittest
```

If a change is meant to change the results, update the expected values with
`python tests/update_expected.py` and describe the change, with before and after
numbers, in [CHANGELOG.md](CHANGELOG.md).

pySTREAM has been tested with Python 3.14 on Windows, with the package versions allowed
by `requirements.txt`. The macOS version has not been tested on a Mac yet.
