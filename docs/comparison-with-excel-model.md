# pySTREAM and the Excel STREAM model

pySTREAM is a Python version of the Excel STREAM model. This note describes how
closely the two agree, where pySTREAM deliberately differs, and some questions about
the Excel model itself, for whoever maintains either of them.

The reference is `STREAM_bioPlus_2015baseyear for validation.xlsm` (base year 2015,
scenario 2050), using the results Excel had last calculated. The Excel STREAM model is
described at https://www.esymodels.man.dtu.dk/stream; the workbooks mentioned in this
note are not included with pySTREAM.

## What is compared, and matches

`tests/test_excel_reference.py` compares pySTREAM's default scenario with the Excel
results (copied, with their cell references, into `tests/excel_reference.py`). These
agree to within a relative difference of 1e-9, for both the base year and the
scenario year:

- total electricity and district heating demand
- primary energy per fuel type
- CO2 emissions per sector and fuel type, with and without CCS
- capital and O&M costs per sector (electricity, district heat, the four sectors,
  fuel refinery capital costs)
- transport capital costs: the vehicles' investment and maintenance (`Transport`
  sheet rows 59 and 168, also checked per fuel), shown in both models as transport
  "Capital Cost" without separate O&M costs
- CO2 costs per tonne of CO2 per sector (the Excel scenario uses a CO2 price of
  32.7 EUR/t, pySTREAM's scenario data 50 EUR/t)
- total fuel costs in the base year (the Excel scenario uses other fuel prices)

The test fails if a change makes pySTREAM differ from these values.

## Where pySTREAM deliberately differs

**Fuel refinery O&M costs.** pySTREAM is 81.5 M EUR higher than Excel in the base
year and 348 M EUR higher in the scenario. In `Fuel Refinery-Tech`, Excel's
"Operation costs, fixed" (columns DM:DO) is `Var. O&M + Transportation Cost`
(columns O and F), and "Operation costs, variable" (DP:DR) is column U, which is the
*electricity input factor*, not a cost. So Excel uses variable O&M + transport + the
electricity factor, and leaves out the fixed O&M. This looks like references that
moved when columns were inserted. pySTREAM uses transport + fixed + variable O&M from
its data. The difference is large in the scenario because biofuels have high fixed
O&M (biodiesel 3.4 EUR/GJ, upgraded biogas 7.2 EUR/GJ).

**CO2 costs of electricity production.** Excel's "CO2 cost for electricity and CHP
production" (`Economic balance`, column K) equals the emissions in Mt; the CO2 price
is not applied (base year 6.37 M EUR for 6.37 Mt at 5 EUR/t). pySTREAM uses
CO2 price x emissions after CCS, as Excel does for all other sectors.

**The hourly time series** is pySTREAM's own simplified version of STREAM's separate
hourly ("duration curve") model, which is not available. The general electricity
consumption is defined exactly as in STREAM's export to that model
(`stream_flow_model_pc.xlsm`, sheet `Exp to Dur.Curve`). The rules for flexible
consumption (electric vehicles, fuel refinery) are taken from comments in the code;
the hourly consumption adds up to the annual demand (tested). Heat pumps and electric
boilers in district heating follow the heat consumption profile, because pySTREAM has
no hourly heat production.

**Electricity and district heating inputs for biogas.** Excel's refinery lists 0.082 GJ
electricity and 0.194 GJ district heat per GJ biogas; pySTREAM's data has 0. They have
no effect in Excel either: its refinery electricity and heat demand are the same as
pySTREAM's, and both are priced at 0.

## Questions about the Excel model

pySTREAM gives the same results as Excel for these, but the Excel calculation looks
questionable. They are not changed, to stay consistent with Excel.

1. **Sector CO2 is divided by the boiler efficiency.** In the sector calculation sheets
   (e.g. `Ind-Calc` column R), CO2 = emission factor x fuel use / efficiency, while
   the fuel costs (column L) use fuel use x price without dividing. If "fuel use" is the
   fuel input, as the demand calculations treat it, CO2 is overstated by 1/efficiency:
   about 1.7 Mt in the base year (5 % of 32.2 Mt) and 0.07 Mt in the scenario.
2. **District heating in the base year is produced 1.5 times.** The base year's district
   heating boilers produce 124.1 PJ (an input) and CHP plants another 54.1 PJ, against a
   demand of 119.0 PJ including grid losses. The boiler input looks like the total
   district heat production including CHP.
3. **The base year and the scenario split transport energy over fuels differently.** In
   the base year, each transport type's energy is split by the fuels' shares of the
   transport work; in the scenario by fuel share x the fuel's own specific energy demand.
   As an electric train uses about a fifth of the energy of a diesel train, the base
   year's transport electricity is 2.36 PJ with the first method and 0.75 PJ with the
   second.
