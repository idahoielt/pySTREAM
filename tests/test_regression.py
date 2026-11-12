"""Regression tests: the model must keep producing the stored results.

The expected results in tests/expected/ were produced by the model itself. If a
test fails, a change has altered the model results. If that change was intended,
run `python tests/update_expected.py` and document the change in docs/CHANGELOG.md.
"""
import os
import shutil
import tempfile
import unittest

import openpyxl

from tests import helpers


class RegressionTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.results = helpers.run_all_scenarios()

    def check_scenario(self, scenario):
        if 'error' in self.results[scenario]:
            self.fail(f'The model run failed: {self.results[scenario]["error"]}')
        differences = helpers.compare(helpers.load_expected(scenario), self.results[scenario])
        if differences:
            self.fail(f"{len(differences)} results differ from the expected values, e.g.\n  "
                      + "\n  ".join(differences[:20]))

    def test_default(self):
        self.check_scenario('default')

    def test_year_2035(self):
        self.check_scenario('year_2035')

    def test_heat_and_transport_electrification(self):
        self.check_scenario('heat_and_transport_electrification')

    def test_biogas_boiler_and_industry_oil(self):
        self.check_scenario('biogas_boiler_and_industry_oil')

    def test_oil_heat_imports_and_gasoline(self):
        self.check_scenario('oil_heat_imports_and_gasoline')

    def test_chp_surplus(self):
        self.check_scenario('chp_surplus')

    def test_no_wind_or_solar(self):
        self.check_scenario('no_wind_or_solar')

    def test_high_prices(self):
        self.check_scenario('high_prices')

    def test_model_runs_from_another_folder(self):
        # Students often start pySTREAM from another folder (e.g. VS Code opened one level up)
        program = helpers.make_program_copy('default')
        elsewhere = tempfile.mkdtemp()
        try:
            results = helpers.run_model(program, cwd=elsewhere)
        finally:
            shutil.rmtree(program, ignore_errors=True)
            shutil.rmtree(elsewhere, ignore_errors=True)
        expected = helpers.load_expected('default')
        for year in ('BY', 'SC'):
            self.assertEqual(helpers.compare(expected[year], results[year]), [])

    def test_electricity_production_equals_demand(self):
        for scenario, results in self.results.items():
            if 'error' in results:
                continue
            sc = results['SC']
            production = sum(row['Amount'] for row in sc['power'].values() if row['Amount'] is not None)
            self.assertAlmostEqual(production, sc['total_electricity'], places=6, msg=f'{scenario}: production vs demand [PJ]')
            negative = {tech: row['Amount'] for tech, row in sc['heat'].items() if (row['Amount'] or 0) < 0}
            self.assertEqual(negative, {}, f'{scenario}: negative district heating boiler production')

    def test_co2_costs_match_emissions(self):
        # CO2 cost = CO2 price x emissions (negative emissions are not paid for).
        # Electricity pays for the CO2 left after CO2 capture and storage (CCS).
        # The refinery pays nothing: refined fuels are paid for where they are used.
        emission_column = {'Electricity': 'Electricity production w. CCS', 'District Heat': 'District Heat',
                           'Residential': 'Residential', 'Tertiary': 'Tertiary',
                           'Industry': 'Industry', 'Agriculture': 'Agriculture', 'Transport': 'Transport'}
        for scenario, results in self.results.items():
            for year in ('BY', 'SC'):
                if 'error' in results:
                    continue
                r = results[year]
                for sector, column in emission_column.items():
                    cost = sum(row[sector] or 0 for row in r['costs']['CO2'].values())
                    emissions = sum(max(row[column] or 0, 0) for row in r['CO2_by_fuel'].values())
                    self.assertAlmostEqual(cost, r['co2_price'] * emissions, places=6, msg=f'{scenario} {year} {sector}')
                refinery_cost = sum(row['Fuel Refinery'] or 0 for row in r['costs']['CO2'].values())
                self.assertEqual(refinery_cost, 0, f'{scenario} {year}: the fuel refinery should not pay CO2 costs')

    def test_fuels_used_in_dh_boilers_have_a_fuel_cost(self):
        for scenario, results in self.results.items():
            for year in ('BY', 'SC'):
                if 'error' in results:
                    continue
                r = results[year]
                used = [fuel for fuel, row in r['fuel_consumption'].items() if (row['DH boilers'] or 0) > 0]
                missing = [fuel for fuel in used if r['costs']['fuel'][fuel]['District Heat'] is None]
                self.assertEqual(missing, [], f'{scenario} {year}: fuels in DH boilers without fuel cost')

    def test_hourly_consumption_adds_up_to_annual_demand(self):
        for scenario, (_, has_time_series) in helpers.SCENARIOS.items():
            results = self.results[scenario]
            if not has_time_series or 'error' in results:
                continue
            hourly = results['time_series']['hourly_sum']
            consumption_MWh = hourly['Consumption after renewables (unflexible and flexible)'] + hourly['Renewables power']
            annual_MWh = results['SC']['total_electricity_demand'] * 1e6 / 3.6
            self.assertAlmostEqual(consumption_MWh / annual_MWh, 1, places=9, msg=f'{scenario}: hourly vs annual electricity demand')

    def test_assumptions_are_read_from_the_assumptions_file(self):
        program = helpers.make_program_copy('default')
        try:
            path = os.path.join(program, 'data', 'Assumptions.xlsx')
            wb = openpyxl.load_workbook(path)
            helpers.set_value(wb['Assumptions'], 'Grid losses, electricity', 'Value', 0.1)
            helpers.set_value(wb['Assumptions'], 'Discount rate', 'Value', 0.06)
            wb.save(path)
            results = helpers.run_model(program)
        finally:
            shutil.rmtree(program, ignore_errors=True)
        sc = results['SC']
        self.assertAlmostEqual(sc['gridloss'] / sc['total_electricity'], 0.1, places=9)
        capital = lambda r: sum(v or 0 for row in r['SC']['costs']['capital'].values() for v in row.values())
        self.assertGreater(capital(results), capital(helpers.load_expected('default')), 'a higher discount rate must raise capital costs')

    def test_scenario_files_without_transport_costs_use_the_defaults(self):
        # scenario files saved before the 'Transport Costs' sheet was added must still work
        program = helpers.make_program_copy('default')
        try:
            path = os.path.join(program, 'data', 'Scenario Data.xlsx')
            wb = openpyxl.load_workbook(path)
            del wb['Transport Costs']
            wb.save(path)
            results = helpers.run_model(program)
        finally:
            shutil.rmtree(program, ignore_errors=True)
        self.assertEqual(helpers.compare(helpers.load_expected('default')['SC'], results['SC']), [])

    def test_time_series_has_no_missing_values(self):
        for scenario, (_, has_time_series) in helpers.SCENARIOS.items():
            if has_time_series and 'error' not in self.results[scenario]:
                hourly_sums = self.results[scenario]['time_series']['hourly_sum']
                missing = [column for column, value in hourly_sums.items() if value is None]
                self.assertEqual(missing, [], f'{scenario}: hourly columns with missing values')


if __name__ == '__main__':
    unittest.main()
