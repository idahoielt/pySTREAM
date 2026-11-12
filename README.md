# pySTREAM

pySTREAM is an energy scenario modelling tool with a simple desktop window. You describe
an energy system for a future year — how electricity and heat are produced, which fuels
the sectors and transport use — and pySTREAM calculates the energy balance, CO2
emissions, costs and an hourly time series, and compares them with the base year.

pySTREAM is a Python version of the Excel model
[STREAM](https://www.esymodels.man.dtu.dk/stream) and gives the same results (see
[docs/comparison-with-excel-model.md](docs/comparison-with-excel-model.md)). It was
created by Ida Cecilie Hoielt as a Bachelor's thesis project at DTU (Technical
University of Denmark) in 2023 with Tilman Weckesser as supervisor. It was updated as part of a special course in 2024, and in 2026 reviewed with assistance from Claude to fix bugs and add some QoL features.

## Install

You need **Python 3.12 or newer**. If you do not have it, install it from
https://www.python.org/downloads/.

1. Download pySTREAM: on GitHub, click **Code → Download ZIP**, and unzip it, for
   example to your Desktop or Documents folder.
2. Install it by double-clicking the install script in the pySTREAM folder:
   - **Windows:** `install_windows.bat`
   - **macOS:** `install_mac.command`. The first time, macOS may refuse to open a
     script downloaded from the internet ("Apple could not verify ..."). Then:
     1. Close the message (do not move the script to the Trash).
     2. Open **System Settings → Privacy & Security**, scroll down to **Security**
        and click **Open Anyway** next to the message about the script.
     3. Confirm with **Open Anyway** (and your Mac password if asked).

     On macOS 14 (Sonoma) or older you can instead right-click (or Ctrl-click) the
     script, choose **Open** and confirm. The same can happen the first time you open
     `start_mac.command`; allow it in the same way.

   This takes a few minutes. It creates a Python environment in the `.venv` folder and
   installs the packages pySTREAM needs (listed in `requirements.txt`).
3. Start pySTREAM by double-clicking `start_windows.bat` (Windows) or
   `start_mac.command` (macOS). The window can take up to a minute to appear.

You only install once; after that, use the start script.

## Your first scenario

pySTREAM starts with a scenario for 2050 compared with the base year 2015. Try changing
it:

1. Click **Edit production** in the menu on the left.
2. In the **Electricity** table, double-click the share of *Coal incl. CO2-storage*,
   type a smaller number, for example `14.6`, and press **Enter**. Then double-click
   *Wind, offshore* and add the same amount, for example `33.4`.
3. Click **Save electricity changes** below the table. The amounts next to the shares
   are updated.
4. Click **Update Scenario** (top right). pySTREAM recalculates everything and saves
   your scenario.
5. Look at the results: **Graphs** (energy flows and more), **CO2 emissions**,
   **Economy** (costs per sector) and **Time series** (hourly electricity).

Good to know:

- An edit only counts after you press **Enter** and then the yellow **Save ... changes**
  button. **Update Scenario** then calculates the new results.
- Shares that must add up to 100 % (heating per sector, fuels per transport type,
  transport work) are checked: pySTREAM tells you which ones do not add up.
- If electricity or district heat production adds up to less than 100 %, the rest is
  imported; above 100 %, the surplus is exported.
- The scenario year (top right) changes the energy demand; press
  **Update Scenario Year** after typing a new year.
- The pySTREAM manual (`pySTREAM Manual.pdf`) explains every page.

## Saving your work

pySTREAM keeps the current scenario in `data/Scenario Data.xlsx`, so it is still there
the next time you start pySTREAM. The buttons at the bottom of the menu:

- **Save scenario as...** saves the current scenario in a file of your choice, so you can
  keep several scenarios.
- **Open scenario...** continues with a saved scenario.
- **Reset to default scenario** goes back to the scenario pySTREAM comes with.

Every plot is also saved as a picture in the `figures` folder, for use in a report.

## If something goes wrong

- pySTREAM shows errors in a window, in plain language. The details are saved in
  `pySTREAM error log.txt` in the pySTREAM folder; include that file if you ask for help.
- **"The file is open in Excel":** close the data file in Excel and try again.
- **The energy flow (Sankey) diagrams** are drawn with Google Chrome. Without Chrome,
  the Graphs page shows a button that opens them in your web browser instead.
- **pySTREAM adapts to your screen:** the window, text, tables and plots are scaled to
  fill the screen pySTREAM is started on. On a very small screen (smaller than about
  1130 x 610 pixels) the text keeps a readable size and scroll bars appear; you can also
  scroll with the mouse wheel (Shift + wheel scrolls sideways). If you move pySTREAM to
  another screen, start it again there to fit that screen.
- **The install script says Python is missing or too old:** install Python 3.12 or newer
  from https://www.python.org/downloads/ and run the install script again.

## Data and assumptions

- `data/Base Year Data.xlsx`: the energy system in the base year 2015. The data is based
  on the STREAM model; updating it to a more recent base year is planned for the future.
- `data/Default Scenario Data.xlsx`: the scenario pySTREAM comes with.
- `data/Energy profiles.xlsx`: hourly profiles of electricity and heat demand and of
  wind, solar, wave and hydro power, for the time series.
- `data/Assumptions.xlsx`: model assumptions such as grid losses, the discount rate and
  the CO2 capture rate, with units and explanations. You can change them there.

How pySTREAM's results compare with the Excel STREAM model, and where they deliberately
differ, is described in [docs/comparison-with-excel-model.md](docs/comparison-with-excel-model.md).
All changes since the original 2023 version are listed in [CHANGELOG.md](CHANGELOG.md).

## For developers

```
pySTREAM/
├── pySTREAM.py           # the window (tkinter); start pySTREAM with this file
├── Backend.py            # the model: energy balance, CO2 and costs
├── Time_Series.py        # the hourly time series
├── data/                 # input data (see above)
├── docs/                 # comparison with the Excel STREAM model
├── tests/                # automatic tests
├── requirements.txt      # Python packages
├── install_*/start_*     # install and start scripts for Windows and macOS
└── figures/              # plots saved while running (created automatically)
```

To set up without the install scripts:

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
source .venv/bin/activate       # macOS/Linux
pip install -r requirements.txt
python pySTREAM.py
```

On Linux, tkinter may need to be installed separately (e.g. `sudo apt install python3-tk`).

### Tests

The tests check that results do not change unintentionally, that they match the Excel
STREAM model, and that the main parts of the window work. Run them from the pySTREAM
folder with the environment active (takes a few minutes):

```bash
python -m unittest
```

If a change is meant to change the results, update the expected values with
`python tests/update_expected.py` and describe the change, with before and after
numbers, in `CHANGELOG.md`.

pySTREAM has been tested with Python 3.14 on Windows, with the package versions allowed
by `requirements.txt`. The macOS scripts have not been tested on a Mac yet.

## How to cite

If you use pySTREAM in your work, please cite it; see [CITATION.cff](CITATION.cff)
(GitHub shows it as **Cite this repository**).

## License

pySTREAM is released under the [MIT License](LICENSE).

## Credits

Created by Ida Cecilie Hoielt as a Bachelor's thesis project at DTU, 2023. Based on the
STREAM energy system model: https://www.esymodels.man.dtu.dk/stream
