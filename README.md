# pySTREAM

pySTREAM is an energy scenario modelling tool with a simple desktop window. You describe
an energy system for a future year — how electricity and heat are produced, which fuels
the sectors and transport use — and pySTREAM calculates the energy balance, CO2
emissions, costs and an hourly time series, and compares them with the base year.

It is a Python version of the Excel model [STREAM](https://www.esymodels.man.dtu.dk/stream)
and gives the same results.

**Contents:** [Install on Windows](#install-on-windows) ·
[Install on macOS](#install-on-macos) · [Install with conda](#install-with-conda-optional) ·
[Your first scenario](#your-first-scenario) ·
[Saving your work](#saving-your-work) · [If something goes wrong](#if-something-goes-wrong) ·
[More information](#more-information)

## Install on Windows

You install pySTREAM once; after that you only use step 5.

1. **Install Python 3.12 or newer** (skip this if you already have it). Download it from
   https://www.python.org/downloads/, run the installer and keep the default options.
2. **Download pySTREAM:** on this GitHub page, click the green **Code** button, then
   **Download ZIP**.
3. **Unzip it:** right-click the downloaded ZIP file, choose **Extract All...** and pick a
   folder, for example `Documents`. Use the extracted folder, not the ZIP file itself.
4. **Install:** open the pySTREAM folder and double-click **`install_windows.bat`**.
   A black window shows the progress; this takes a few minutes. When it says that
   pySTREAM is installed, press a key to close it.
5. **Start pySTREAM:** double-click **`start_windows.bat`**. The window can take up to a
   minute to appear. Keep the black window open while you use pySTREAM.

If Windows shows a security warning when you double-click a script (it was downloaded
from the internet), choose **Run**, or click **More info** and then **Run anyway**.

## Install on macOS

You install pySTREAM once; after that you only use step 5.

1. **Install Python 3.12 or newer** (skip this if you already have it). Download the macOS
   installer from https://www.python.org/downloads/, open it and follow the steps.
2. **Download pySTREAM:** on this GitHub page, click the green **Code** button, then
   **Download ZIP**.
3. **Unzip it:** double-click the downloaded ZIP file (usually in `Downloads`). Move the
   pySTREAM folder that appears to where you want to keep it, for example `Documents`.
4. **Install:** open the pySTREAM folder and double-click **`install_mac.command`**.
   A Terminal window shows the progress; this takes a few minutes. When it says that
   pySTREAM is installed, press Enter and close the window.
5. **Start pySTREAM:** double-click **`start_mac.command`**. The window can take up to a
   minute to appear. Keep the Terminal window open while you use pySTREAM.

**If macOS does not open a script** ("Apple could not verify ..."), which is normal for a
script downloaded from the internet the first time:

1. Close the message. Do not move the script to the Trash.
2. Open **System Settings → Privacy & Security**, scroll down to **Security** and click
   **Open Anyway** next to the message about the script.
3. Confirm with **Open Anyway** (and your Mac password if asked).

On macOS 14 (Sonoma) or older you can instead right-click (or Ctrl-click) the script,
choose **Open** and confirm. This can happen for both scripts, the first time you open each.

## Install with conda (optional)

If you already use Anaconda or Miniconda, you can install pySTREAM with conda instead of
the install script. Python is then installed by conda, so you can skip step 1 above.

1. **Download and unzip pySTREAM** as in steps 2 and 3 above.
2. **Open a terminal in the pySTREAM folder:** on Windows, open the **Anaconda Prompt**;
   on macOS, open **Terminal**. Then go to the folder with `cd`, for example:

   ```bash
   cd Documents/pySTREAM-main
   ```

3. **Install** (once; this takes a few minutes). It creates a conda environment called
   `pystream` with Python and the packages pySTREAM needs:

   ```bash
   conda env create -f environment.yml
   ```

4. **Start pySTREAM** (every time, from the pySTREAM folder):

   ```bash
   conda activate pystream
   python program/pySTREAM.py          # Windows
   python program/pySTREAM_mac.py      # macOS
   ```

   Keep the terminal open while you use pySTREAM. The start scripts
   (`start_windows.bat`, `start_mac.command`) do not work with conda.

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
- The [pySTREAM manual](pySTREAM%20Manual.pdf) (`pySTREAM Manual.pdf` in the pySTREAM
  folder) explains every page.

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
- **The install script says Python is missing or too old:** install Python 3.12 or newer
  from https://www.python.org/downloads/ and run the install script again.
- **"pySTREAM is not installed yet":** run the install script (step 4) first.
- **"The file is open in Excel":** close the data file in Excel and try again.
- **The energy flow (Sankey) diagrams** are drawn with Google Chrome. Without Chrome,
  the Graphs page shows a button that opens them in your web browser instead.
- **The window and text size:** pySTREAM scales itself to fill the screen it is started
  on. On a very small screen (smaller than about 1130 x 610 pixels) scroll bars appear;
  you can also scroll with the mouse wheel (Shift + wheel scrolls sideways). If you move
  pySTREAM to another screen, start it again there to fit that screen.
- **Installing again:** delete the `.venv` folder in the pySTREAM folder (it may be hidden)
  and run the install script again.

## More information

- **What is in the pySTREAM folder:** the install and start scripts, the
  [manual](pySTREAM%20Manual.pdf), `data` (the input data and your scenario),
  `figures` (your plots), `program` (the Python files; you do not need to open them),
  `docs` and `tests` (for developers).
- **Data and assumptions:**
  - `data/Base Year Data.xlsx`: the energy system in the base year 2015. The data is
    based on the STREAM model; updating it to a more recent base year is planned.
  - `data/Default Scenario Data.xlsx`: the scenario pySTREAM comes with.
  - `data/Energy profiles.xlsx`: hourly profiles of electricity and heat demand and of
    wind, solar, wave and hydro power, for the time series.
  - `data/Assumptions.xlsx`: model assumptions such as grid losses, the discount rate and
    the CO2 capture rate, with units and explanations. You can change them there.
- **Comparison with the Excel STREAM model:** how the results compare, and where they
  deliberately differ: [docs/comparison-with-excel-model.md](docs/comparison-with-excel-model.md).
- **Changes** since the original 2023 version: [docs/CHANGELOG.md](docs/CHANGELOG.md).
- **For developers** (folders, setting up by hand, tests):
  [docs/developers.md](docs/developers.md).

## How to cite

If you use pySTREAM in your work, please cite it; see [CITATION.cff](CITATION.cff)
(GitHub shows it as **Cite this repository**).

## License

pySTREAM is released under the [MIT License](LICENSE).

## Credits

pySTREAM was created by Ida Cecilie Hoielt as a Bachelor's thesis project at DTU (Technical
University of Denmark) in 2023 with Tilman Weckesser as supervisor. It was updated as part
of a special course in 2024, and in 2026 reviewed with assistance from Claude to fix bugs
and add some QoL features. It is based on the STREAM energy system model:
https://www.esymodels.man.dtu.dk/stream
