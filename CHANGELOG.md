# Changelog

Changes since the original version of pySTREAM (git tag `original-version`).

## Unreleased

### Changes to model results

Every change listed here was made on purpose. The numbers refer to the test
scenarios in `tests/` ("default" is the data shipped with pySTREAM).

- **Scenario transport demand is no longer rounded down to whole numbers.**
  The Excel model does not round it. The scenario year now matches the Excel
  model to about 1e-12. Effect on the default scenario (base year unchanged):

  | Result (scenario year) | Before | After |
  |---|---|---|
  | Electricity demand [PJ] | 124.893131 | 124.893123 |
  | District heating demand [PJ] | 103.329491 | 103.329211 |
  | Primary energy, bioenergy [PJ] | 283.602458 | 283.604144 |
  | CO2 emissions incl. CCS [Mt] | -1.033283 | -1.033298 |
  | Total of all costs [M EUR] | 7825.68 | 7825.71 |

- **The fuel refinery now produces the biogas burned in district heating
  boilers.** Before, only biogas for power plants was produced, so the manure
  and energy crops for the boilers' biogas, and the related refinery costs and
  CO2, were missing. Matches the Excel model now. Base year (2.04 PJ biogas in
  boilers; the default scenario has none, so only its base year changes):

  | Result (base year) | Before | After | Excel |
  |---|---|---|---|
  | Primary energy, bioenergy [PJ] | 134.190 | 138.064 | 138.064 |
  | Fuel refinery CO2, bioenergy [Mt] | -0.621 | -0.927 | -0.927 |
  | Fuel refinery capital costs [M EUR] | 722.39 | 734.96 | 734.96 |
  | CO2 emissions incl. CCS [Mt] | 32.378 | 32.072 | |
  | Capital / O&M / fuel costs [M EUR] | 2429.4 / 1560.5 / 3687.7 | 2441.9 / 1579.9 / 3655.8 | |

  The total cost hardly changes: about 32 M EUR moves from fuel costs (the
  boilers paid the full refined biogas price) to the refinery's capital and O&M
  costs.

- **Industry oil boilers are treated as using diesel everywhere.** The fuel
  refinery already made diesel for them, but CO2, fuel costs and the Sankey
  diagram treated them as burning crude oil: CO2 used the oil emission factor,
  and industry paid the crude oil price although the refinery needs 1.51 GJ of
  crude per GJ of diesel. Base year (16.6 PJ; none in the default scenario):

  | Result (base year) | Before | After | Excel |
  |---|---|---|---|
  | Industry CO2 [Mt] | 1.5375 | 1.6206 | 1.6206 |
  | Industry CO2 costs [M EUR] | 19.57 | 19.99 | 19.99 |
  | Fuel costs, all sectors [M EUR] | 3655.77 | 3895.77 | |

  In the test scenario with 20 % oil boilers in industry, scenario-year fuel
  costs rise by 203 M EUR and CO2 by 0.055 Mt.

- **Electricity production now equals demand when CHP heat exceeds the district
  heating demand.** In that case the district heating boilers are switched off,
  but power production had been calculated with negative boiler production
  (heat pumps "giving back" electricity) and was not recalculated, so less
  electricity was produced than needed. This happens easily, e.g. when
  residential heating moves from district heating to heat pumps. "Update
  Scenario" now also warns that the surplus heat is not used. The default
  scenario is not affected. Scenario year of the affected test scenarios:

  | Test scenario | Electricity production before -> after [PJ] (demand) | Primary energy [PJ] | All costs [M EUR] |
  |---|---|---|---|
  | 20 % residential heat to heat pumps, electric cars | 139.06 -> 140.43 (140.43) | 504.7 -> 507.7 | 7310 -> 7351 |
  | Much coal CHP | 121.81 -> 124.38 (124.38) | 630.3 -> 638.0 | 10435 -> 10569 |
  | No wind or solar | 123.73 -> 124.38 (124.38) | 584.8 -> 586.5 | 9562 -> 9592 |

- **CO2 costs of electricity production take CO2 capture and storage (CCS) into
  account.** Before, plants with CCS paid for all the CO2 they captured, so CCS
  brought no saving. The Excel model uses the emissions after CCS. Default
  scenario: electricity CO2 costs 567.8 -> 87.1 M EUR, all CO2 costs
  712.0 -> 231.3 M EUR. The base year has no CCS and is unchanged.

- **The fuel refinery no longer pays CO2 costs for CO2 that consumers already
  pay.** The refinery paid for crude oil's CO2 while transport and industry paid
  for the same CO2 as diesel and gasoline. As in the Excel model, refined fuels'
  CO2 is now paid only where they are used. Default scenario, all CO2 costs:
  base year 218.5 -> 146.3 M EUR, scenario year 231.3 -> 178.0 M EUR. CO2
  emissions are unchanged.

- **Industry boilers have variable O&M costs, as in the Excel model**
  (0.75 EUR/GJ for wood pellet and coal boilers, 0.1389 EUR/GJ for electric
  heaters; new column in the "Sectors Details" sheet of both data files).
  Default scenario, industry O&M costs: base year 11.72 -> 19.95 M EUR,
  scenario year 6.77 -> 31.82 M EUR (both equal to Excel).

- **Fuel costs of district heating boilers using fuels without CHP are
  counted.** Costs of fuels only used in boilers (oil, heat imports, ...)
  disappeared because of an empty value in a sum. Default scenario: base year
  district heat fuel costs +17.62 M EUR (oil boilers). Heat imports were always
  free; in the test scenario with 10 % heat imports and 10 % oil boilers they now
  cost 6.38 and 10.11 M EUR in the scenario year.

- **Gasoline includes 1 % blended ethanol, as in the Excel model.** Refining
  1 GJ of gasoline also uses 0.01 GJ of ethanol at its market price (new column
  in the refinery input sheet of both data files). It counts in the cost only,
  like in Excel. Base year gasoline costs 976.48 -> 979.93 M EUR; the default
  scenario has no gasoline. The base year's total fuel costs now equal the
  Excel model's (3916.84 M EUR).

- **Transport costs include the vehicles, as in the Excel model.** pySTREAM
  left the transport capital costs empty (the Economy page showed 0.0). The
  Excel model calculates the vehicles' investment and maintenance per transport
  type and fuel: transport work x cost per person km or tonne km (new "Transport
  Costs" sheet in the scenario data; the base year sheet already existed but was
  not used). They are shown as transport capital costs; there are no separate
  transport O&M costs. Scenario files saved before this change use the default
  scenario's costs per km. Default scenario, equal to Excel per fuel:

  | Result | Base year | Scenario year |
  |---|---|---|
  | Transport capital costs [M EUR] | 0 -> 2455.42 | 0 -> 5797.41 |
  | Capital costs, all sectors [M EUR] | 2441.92 -> 4897.34 | 2121.11 -> 7918.52 |
  | Total of all costs [M EUR] | 8093.13 -> 10548.55 | 7316.76 -> 13114.17 |

  Vehicles are now the largest cost item. The costs depend on the transport
  demand and on the fuel shares (in the data, a car costs 0.090 EUR per person
  km with electricity, 0.036 with diesel and 0.101 with biodiesel).

- **The hourly electricity consumption adds up to the annual demand.** Part of
  the electric vehicle charging was subtracted instead of added, fuel refinery
  electricity was counted wrongly, and electricity for district heating heat
  pumps and electric boilers was missing. Only the time series changes:

  | Time series | Hourly consumption [TWh] (annual demand) | Imports / exports [TWh] |
  |---|---|---|
  | Default scenario | 34.837 -> 34.693 (34.693) | 1.088 / 0.944 -> 0.962 / 0.962 |
  | Electrification test scenario | 37.086 -> 39.009 (39.009) | 0.004 / 1.214 -> 0.940 / 0.940 |
  | No wind or solar test scenario | 34.826 -> 34.551 (34.551) | 0.273 / 0 -> 0.010 / 0 |

  Plant capacities and operating hours change accordingly (e.g. biogas in the
  electrification scenario 1993 -> 891 MW).

- **Time series works in scenarios without surplus renewable power.** Before,
  such a scenario (e.g. no wind or solar) made the time series run forever,
  and pySTREAM could not start. The electricity use that is normally placed in
  hours with surplus wind/solar is now spread evenly over the year in that case.
  Only affects scenarios that used to hang; there were no results before.

### Fixes (model results unchanged)

- The CO2 page showed about three times the real emissions (base year: 97.13 Mt
  shown, 32.38 Mt calculated). It now shows the calculated totals.
- "Update Scenario" did not recalculate capital and O&M costs, so the Economy
  page kept showing the costs of the scenario pySTREAM was started with.
- The time series used the electricity demand of the scenario pySTREAM was
  started with, even after "Update Scenario".
- The "Update Time Series" button crashed after recalculating.
- The "Update Scenario Year" button crashed on every use (with pandas 3), after
  already saving the new year to the data file. Invalid years are now refused
  without changing anything.
- Saving the electricity table turned the electricity import amount into an
  empty value (NaN).
- Import shares calculated as "100 % minus the other shares" are rounded, so
  floating-point rest values (like 0.0000000000000001) are no longer treated as
  real imports in the time series.
- pySTREAM can be started from any folder: data and images are found relative
  to the program folder ("No such file or directory" error in the manual).
- Plots are saved in a new `figures` folder instead of `images`, so running
  pySTREAM no longer changes files under version control.
- pySTREAM no longer stops at startup on macOS and Linux. It loaded a
  Windows-only window icon; the icon has been removed.
- "Update Scenario" warns when CHP plants produce more district heat than is
  needed (see above).
- The Sankey diagrams no longer stop or freeze pySTREAM. Their images are made
  with Kaleido, which needs Google Chrome and sometimes hangs. If making an image
  fails or takes more than 60 seconds, the Graphs page shows a button that opens
  the diagram in the browser instead (the manual's "small window that does not
  show anything" problem).
- Closing the window now always ends pySTREAM; Kaleido's background threads
  sometimes kept it running.
- The Sankey diagram left out coal used in CHP plants for heat (20.46 PJ in the
  default scenario); all flows are now drawn. The duration curve's fixed y-axis
  no longer cuts off part of the curve, and resource bars that are cut off at
  500 PJ show their value.

### Usability

- macOS: the buttons keep their colours. On macOS, Tk draws its standard
  buttons as Mac buttons and ignores their colour, which made the white text of
  the menu buttons hard to read; pySTREAM now uses its own buttons, which look
  the same on every system.
- macOS: plots are drawn with Tk, like the rest of the window. matplotlib
  would otherwise use Apple's own window system there, and running two window
  systems in one program can make it freeze or crash.
- pySTREAM adapts to the screen it is started on: the whole layout, text,
  tables and plots are scaled to fill it, from small laptops to 4K screens
  (the layout is designed for 1920 x 1080 pixels). On Windows with display
  scaling above 100 %, text is now sharp (Windows no longer enlarges the window
  as a picture) and the layout no longer changes after the first plot. On very
  small screens the text keeps a readable minimum size (60 %) and scrollbars
  appear; the mouse wheel scrolls the page. On macOS, text is now scaled in the
  same way as on Windows (before, it was smaller there). No model results
  change.
- Errors are shown in a window in plain language (e.g. "the file is open in
  Excel, close it and try again"), with the details saved in
  `pySTREAM error log.txt`.
- Entered values are checked: out-of-range values are refused with a message
  that explains the allowed range and unit, a decimal comma is accepted, and
  "Update Scenario" is refused while shares that must add up to 100 % do not.
  Electricity and district heat shares above 100 % are shown as exports.
- "Save scenario as", "Open scenario" and "Reset to default scenario" buttons.
- pySTREAM starts faster (the time series is calculated when its page is
  opened), shows a message while starting, and shows a "please wait" message
  during long calculations.

### Added

- `data/Assumptions.xlsx`: model assumptions that were written in the code (grid
  losses, condensing share of extraction plants, discount rate, base year CO2
  price, CO2 capture rate, time series shares) with units and explanations.
  Same values as before, so results are unchanged.
- `data/Default Scenario Data.xlsx`: an unchanged copy of the shipped scenario
  (pySTREAM saves the user's changes into `Scenario Data.xlsx`).
- Double-click scripts to install and start pySTREAM (`install_windows.bat`,
  `start_windows.bat`, `install_mac.command`, `start_mac.command`). The README
  now states the correct Python version (3.12 or newer).
- Automatic tests (`python -m unittest`): regression tests for 8 scenarios, a
  comparison with the Excel STREAM model, and tests of the GUI.
- `docs/comparison-with-excel-model.md`: what is compared with the Excel STREAM
  model, where pySTREAM deliberately differs, and questions about the Excel model.
- A new README for students and developers: installing, a first scenario step by
  step, saving scenarios, what to do if something goes wrong, the data, and tests.
- `LICENSE` (MIT) and `CITATION.cff` (how to cite pySTREAM).

### Removed

- The Excel STREAM workbook used for validation is no longer included; the
  documentation links to STREAM (https://www.esymodels.man.dtu.dk/stream)
  instead. The values the tests compare with are kept in `tests/excel_reference.py`.
- The window icon.
- Links in the data files to other workbooks (with a local folder path); they
  were not used. Values are unchanged.
- The installation and troubleshooting pages of `pySTREAM Manual.pdf`, which
  described the old version (e.g. changing code lines); the README replaces them.
