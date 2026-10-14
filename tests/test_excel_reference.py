"""Checks that pySTREAM gives the same results as the Excel STREAM model.

Compares the default scenario (the data files as shipped) with values copied from
the Excel model (see excel_reference.py). Differences that are known but not yet
fixed are listed in KNOWN_DIFFERENCES. When a fix makes one of them match, the
test fails until the entry is removed, so the list always reflects reality.
"""
import math
import unittest

from tests import helpers
from tests.excel_reference import EXCEL, EXCEL_CO2_PRICE, EXCEL_FUEL_COSTS_TOTAL_BY

# Allowed relative difference per year.
TOLERANCE = {'BY': 1e-9, 'SC': 1e-9}

# Differences from Excel that are not fixed yet (plan item in brackets).
KNOWN_DIFFERENCES = set()


def column_sum(table, column):
    return sum(row[column] for row in table.values() if row.get(column) is not None)


def comparisons(results):
    """Yield (name, year, pySTREAM value, Excel value) for every compared number."""
    for year in ('BY', 'SC'):
        py, xl = results[year], EXCEL[year]
        yield f'{year} total electricity demand', year, py['total_electricity'], xl['total_electricity']
        yield f'{year} total DH demand', year, py['total_DH_demand'], xl['total_DH_demand']
        for fuel_type, value in xl['primary_fuel'].items():
            yield f'{year} primary fuel {fuel_type}', year, py['primary_fuel'][fuel_type]['Total primary fuel consumption'], value
        for column, by_type in xl['CO2'].items():
            for fuel_type, value in by_type.items():
                yield f'{year} CO2 {column} {fuel_type}', year, py['CO2'][fuel_type][column], value
        for sector, value in xl['costs']['capital'].items():
            column = 'Transportation' if sector == 'Transport' else sector  # pySTREAM's name for the column
            yield f'{year} capital costs {sector}', year, column_sum(py['costs']['capital'], column), value
        for sector, value in xl['costs']['O&M'].items():
            yield f'{year} O&M costs {sector}', year, column_sum(py['costs']['O&M'], sector), value
        if year == 'BY':
            total = sum(v or 0 for row in py['costs']['fuel'].values() for v in row.values())
            yield f'{year} total fuel costs', year, total, EXCEL_FUEL_COSTS_TOTAL_BY
        for sector, value in xl['costs']['CO2'].items():
            yield (f'{year} CO2 cost per tonne {sector}', year,
                   column_sum(py['costs']['CO2'], sector) / py['co2_price'], value / EXCEL_CO2_PRICE[year])


class ExcelReferenceTest(unittest.TestCase):

    def test_default_scenario_matches_excel(self):
        results = helpers.run_all_scenarios()['default']
        mismatches = {}
        for name, year, py_value, xl_value in comparisons(results):
            if not math.isclose(py_value, xl_value, rel_tol=TOLERANCE[year], abs_tol=1e-9):
                mismatches[name] = f'{name}: pySTREAM {py_value!r}, Excel {xl_value!r}'
        unexpected = [text for name, text in mismatches.items() if name not in KNOWN_DIFFERENCES]
        now_matching = KNOWN_DIFFERENCES - set(mismatches)
        self.assertEqual(unexpected, [], 'pySTREAM differs from the Excel model')
        self.assertEqual(now_matching, set(), 'These now match Excel; remove them from KNOWN_DIFFERENCES')


if __name__ == '__main__':
    unittest.main()
