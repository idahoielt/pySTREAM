"""Shared helpers for the pySTREAM tests.

The model reads its input files and calculates everything when Backend.py is
imported, and keeps the results in module-level variables. To test a scenario we
therefore copy the program and its data to a temporary folder, change the
scenario workbook in that copy, and run the model in a separate Python process
(see run_model.py). This way the tests never touch the real data files.
"""
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache

import openpyxl

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(TESTS_DIR)
EXPECTED_DIR = os.path.join(TESTS_DIR, 'expected')
PROGRAM_FILES = ['program', 'data']


# Scenario definitions
# ==============================================================================
# Each scenario is a function that edits the copied 'Scenario Data.xlsx' workbook,
# the same way the GUI saves a student's changes.

def find_row(ws, row_label):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 1).value == row_label:
            return row
    raise KeyError(f"Row {row_label!r} not found in sheet {ws.title!r}")


def find_column(ws, column_label):
    for column in range(1, ws.max_column + 1):
        if ws.cell(1, column).value == column_label:
            return column
    raise KeyError(f"Column {column_label!r} not found in sheet {ws.title!r}")


def set_value(ws, row_label, column_label, value):
    ws.cell(find_row(ws, row_label), find_column(ws, column_label), value)


def add_value(ws, row_label, column_label, change):
    cell = ws.cell(find_row(ws, row_label), find_column(ws, column_label))
    cell.value = cell.value + change


def default(wb):
    pass


def year_2035(wb):
    wb['Scenario Year'].cell(2, 1, 2035)


def heat_and_transport_electrification(wb):
    fec = wb['Final Energy Consumption']
    add_value(fec, 'District Heating', 'Residential', -0.20)
    add_value(fec, 'Heatpumps', 'Residential', +0.20)
    fuels = wb['Fuels Transportation']
    for fuel, share in {'Electricity': 0.8, 'Ethanol': 0.1, 'Biodiesel': 0.1}.items():
        set_value(fuels, 'Car', fuel, share)
    for fuel, share in {'Electricity': 0.5, 'Biodiesel': 0.35, 'UPGR Biogas': 0.15}.items():
        set_value(fuels, 'Bus', fuel, share)


def biogas_boiler_and_industry_oil(wb):
    dh = wb['District Heat Boiler Production']
    add_value(dh, 'Wood pellet boiler', 'Shares', -0.20)
    add_value(dh, 'Biogas Boiler', 'Shares', +0.20)
    fec = wb['Final Energy Consumption']
    add_value(fec, 'Wood Pellet Boiler', 'Industry', -0.20)
    add_value(fec, 'Oil Boiler', 'Industry', +0.20)


def oil_heat_imports_and_gasoline(wb):
    # District heating with oil boilers and heat imports (fuels without CHP plants), and gasoline cars
    dh = wb['District Heat Boiler Production']
    add_value(dh, 'Wood pellet boiler', 'Shares', -0.20)
    add_value(dh, 'Oil boiler', 'Shares', +0.10)
    add_value(dh, 'Heat imports', 'Shares', +0.10)
    fuels = wb['Fuels Transportation']
    add_value(fuels, 'Car', 'Ethanol', -0.30)
    add_value(fuels, 'Car', 'Gasoline', +0.30)


def chp_surplus(wb):
    # So much coal CHP that CHP heat alone exceeds the district heating demand
    power = wb['Electricity Production']
    add_value(power, 'Coal incl. CO2-storage', 'Shares', +0.40)
    add_value(power, 'Wind, onshore', 'Shares', -0.244)
    add_value(power, 'Wind, offshore', 'Shares', -0.156)


def no_wind_or_solar(wb):
    power = wb['Electricity Production']
    for tech in ['Wind, offshore', 'Wind, onshore', 'Photo voltaic']:
        add_value(power, 'Natural gas incl. CO2-storage', 'Shares', power.cell(find_row(power, tech), 2).value)
        set_value(power, tech, 'Shares', 0)


def high_prices(wb):
    wb['CO2 Price'].cell(2, 1, 150)
    prices = wb['Fuel Price']
    for fuel, price in {'Natural Gas': 15, 'Biomass (Straw,woodwaste)': 12, 'Coal': 20}.items():
        set_value(prices, fuel, 'Market price [€/GJ]', price)


# name: (edit function, also run the hourly time series?)
SCENARIOS = {
    'default': (default, True),
    'year_2035': (year_2035, False),
    'heat_and_transport_electrification': (heat_and_transport_electrification, True),
    'biogas_boiler_and_industry_oil': (biogas_boiler_and_industry_oil, False),
    'oil_heat_imports_and_gasoline': (oil_heat_imports_and_gasoline, False),
    'chp_surplus': (chp_surplus, False),
    'no_wind_or_solar': (no_wind_or_solar, True),
    'high_prices': (high_prices, False),
}


# Running the model
# ==============================================================================

def make_program_copy(scenario=None):
    """Copy the program and its data to a new temporary folder and apply a scenario.
    The scenario always starts from the shipped 'Default Scenario Data.xlsx', not from
    'Scenario Data.xlsx', which pySTREAM overwrites with the user's own changes."""
    folder = tempfile.mkdtemp(prefix='pystream_test_')
    for name in PROGRAM_FILES:
        source = os.path.join(PROJECT_DIR, name)
        if os.path.isdir(source):
            shutil.copytree(source, os.path.join(folder, name), ignore=shutil.ignore_patterns('Scenario Data.xlsx', '~$*', '__pycache__'))
        else:
            shutil.copy2(source, folder)
    path = os.path.join(folder, 'data', 'Scenario Data.xlsx')
    shutil.copy2(os.path.join(PROJECT_DIR, 'data', 'Default Scenario Data.xlsx'), path)
    if scenario is not None:
        wb = openpyxl.load_workbook(path)
        SCENARIOS[scenario][0](wb)
        wb.save(path)
    return folder


def run_model(program_dir, time_series=False, cwd=None, timeout=600):
    """Run the model in a separate process and return its results as a dict."""
    output_file = os.path.join(program_dir, 'test_results.json')
    command = [sys.executable, os.path.join(TESTS_DIR, 'run_model.py'), program_dir, output_file]
    if time_series:
        command.append('--time-series')
    if cwd:
        command += ['--cwd', cwd]
    completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout)
    if completed.returncode != 0:
        raise RuntimeError(f"Model run failed:\n{completed.stderr[-3000:]}")
    with open(output_file, encoding='utf-8') as f:
        return json.load(f)


@lru_cache(maxsize=None)
def run_all_scenarios():
    """Run every scenario (in parallel) and return {name: results}.
    A scenario that crashes or takes too long gets {'error': message}."""
    def run(name):
        folder = make_program_copy(name)
        try:
            return name, run_model(folder, time_series=SCENARIOS[name][1], timeout=180)
        except Exception as e:
            return name, {'error': f'{type(e).__name__}: {e}'}
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    with ThreadPoolExecutor(max_workers=4) as pool:
        return dict(pool.map(run, SCENARIOS))


GUI_CHECKS = ['co2_totals', 'update_matches_fresh_start', 'update_time_series_button', 'update_scenario_year',
              'start_from_other_folder', 'chp_surplus_warning',
              'sankey_image_fails', 'sankey_image_hangs', 'close_window', 'economy_page', 'plots',
              'file_open_in_excel', 'missing_data_file', 'input_validation',
              'scenario_files', 'time_series_on_demand', 'small_screen',
              'screen_laptop', 'screen_design', 'screen_large', 'screen_tiny']


@lru_cache(maxsize=None)
def run_gui_checks():
    """Run every check in gui_checks.py (in parallel), each in a fresh copy of the program."""
    def run(check):
        folder = make_program_copy()
        output_file = os.path.join(folder, 'gui_check.json')
        try:
            completed = subprocess.run([sys.executable, os.path.join(TESTS_DIR, 'gui_checks.py'), folder, check, output_file],
                                       capture_output=True, text=True, timeout=300)
            if completed.returncode != 0:
                return check, {'error': completed.stderr[-3000:]}
            with open(output_file, encoding='utf-8') as f:
                return check, json.load(f)
        except Exception as e:
            return check, {'error': f'{type(e).__name__}: {e}'}
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    with ThreadPoolExecutor(max_workers=4) as pool:
        return dict(pool.map(run, GUI_CHECKS))


# Collecting results
# ==============================================================================

def number(value):
    """Convert a value to something JSON can store: NaN becomes None, infinity a string."""
    if value is None:
        return None
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(value):
        return None
    if math.isinf(value):
        return 'inf' if value > 0 else '-inf'
    return value


def table(data):
    """DataFrame -> {row: {column: value}}, Series -> {row: value}."""
    if hasattr(data, 'columns'):
        return {str(row): {str(col): number(v) for col, v in values.items()} for row, values in data.iterrows()}
    return {str(row): number(v) for row, v in data.items()}


def collect_outputs(model):
    """Collect the main results from a module that holds pySTREAM's results
    (Backend after import, or the GUI module after 'Update Scenario')."""
    out = {}
    for year in ('BY', 'SC'):
        get = lambda name: getattr(model, f'{name}_{year}')
        out[year] = {
            'total_electricity': number(get('total_electricity_consumption')),
            'gridloss': number(get('gridloss')),
            'total_DH_demand': number(get('total_DH_demand')),
            'el_transport': number(get('el_transport')),
            'el_DH': number(get('el_DH')),
            'el_FR': number(get('el_FR')),
            'co2_price': number(get('co2_price')),
            'transport_consumption': table(get('df_transport_consumption')),
            'transport_fuels': table(get('df_transport_consumption_fuels')),
            'power': table(get('df_Power')[['Amount', 'CHP heat production', 'Total fuel consumption']]),
            'heat': table(get('df_Heat')[['Amount']]),
            'fuel_secondary': table(get('df_fuel_secondary_consumption')),
            'primary_fuel': table(get('df_fuelType_consumption')),
            'fuel_consumption': table(get('df_Fc')),
            'CO2_by_fuel': table(get('df_fuel_CO2')),
            'CO2': table(get('df_fuelType_CO2')),
            'costs': {
                'capital': table(get('df_capital_costs')),
                'O&M': table(get('df_om_costs')),
                'fuel': table(get('df_fuel_costs')),
                'CO2': table(get('df_co2_costs')),
            },
        }
    out['SC']['FEC'] = table(model.df_FEC_SC)
    out['SC']['total_electricity_demand'] = number(model.total_electricity_demand_SC)
    return out


def collect_time_series(df_hour, technologies_used):
    """Summarise the hourly results (sum, min and max of every column)."""
    hourly = df_hour.loc[:, ~df_hour.columns.duplicated()]
    hourly = hourly.drop(columns=['HourOfWeek', 'HourOfMonth'], errors='ignore').astype(float)
    if 'Exported electricity' in hourly:
        hourly['Exported electricity'] = hourly['Exported electricity'].abs()  # the plots flip its sign
    return {
        'technologies': table(technologies_used[['Amount', 'Capacity', 'Hours', 'Avg']]),
        'hourly_sum': table(hourly.sum()),
        'hourly_min': table(hourly.min()),
        'hourly_max': table(hourly.max()),
    }


# Comparing results
# ==============================================================================

def compare(expected, actual, rel_tol=1e-9, abs_tol=1e-9, path=''):
    """Return a list of human-readable differences between two result dicts."""
    if isinstance(expected, dict) and isinstance(actual, dict):
        differences = []
        for key in sorted(set(expected) | set(actual)):
            where = f'{path}/{key}'
            if key not in actual:
                differences.append(f'{where}: missing (expected {expected[key]!r:.60})')
            elif key not in expected:
                differences.append(f'{where}: new value {actual[key]!r:.60}')
            else:
                differences += compare(expected[key], actual[key], rel_tol, abs_tol, where)
        return differences
    if isinstance(expected, float) and isinstance(actual, float):
        if math.isclose(expected, actual, rel_tol=rel_tol, abs_tol=abs_tol):
            return []
    elif expected == actual:
        return []
    return [f'{path}: expected {expected!r}, got {actual!r}']


def expected_file(scenario):
    return os.path.join(EXPECTED_DIR, f'{scenario}.json')


def load_expected(scenario):
    with open(expected_file(scenario), encoding='utf-8') as f:
        return json.load(f)


def save_expected(scenario, results):
    os.makedirs(EXPECTED_DIR, exist_ok=True)
    with open(expected_file(scenario), 'w', encoding='utf-8', newline='\n') as f:
        json.dump(results, f, indent=1, sort_keys=True, ensure_ascii=False)
        f.write('\n')
