"""Tests of what the GUI shows and does. Each check runs the real GUI module
(without showing a window) in a copy of the program; see gui_checks.py."""
import unittest

from tests import helpers


class GuiTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.results = helpers.run_gui_checks()

    def get(self, check):
        result = self.results[check]
        if 'error' in result:
            self.fail(f'The GUI check crashed:\n{result["error"]}')
        return result

    def test_co2_page_shows_model_totals(self):
        result = self.get('co2_totals')
        for year in ('BY', 'SC'):
            for shown, model in zip(result[year]['shown_total'], result[year]['model_total']):
                self.assertAlmostEqual(shown, model, places=2, msg=f'{year} total emissions')
            for fuel, shown_values in result[year]['shown_rows'].items():
                model_values = [v or 0.0 for v in result[year]['model_rows'][fuel]]
                for shown, model in zip(shown_values, model_values):
                    self.assertAlmostEqual(shown, model, places=2, msg=f'{year} emissions of {fuel}')

    def test_gui_works_when_started_from_another_folder(self):
        result = self.get('start_from_other_folder')
        self.assertTrue(result['data_file_saved'], 'Update Scenario should save the data file in the program folder')
        self.assertEqual(result['files_in_start_folder'], [], 'nothing should be written outside the program folder')

    def test_plots_are_saved_in_figures_folder(self):
        result = self.get('start_from_other_folder')
        self.assertIn('sankeySC.png', result['figures_folder'] or [])

    def test_warning_when_chp_heat_exceeds_demand(self):
        messages = self.get('chp_surplus_warning')['messages']
        self.assertTrue(any(kind == 'showwarning' and 'surplus' in text.lower() for kind, title, text in messages),
                        f'expected a warning about surplus heat, got {messages}')

    def test_economy_page_shows_transport_vehicle_costs(self):
        # As in the Excel model, the vehicles' investment and maintenance are shown together as capital costs
        result = self.get('economy_page')
        for year in ('BY', 'SC'):
            model = result[year + ' transport capital costs']
            self.assertGreater(model, 0, f'{year}: vehicle costs missing in the model')
            self.assertEqual(result[year]['Capital costs'][-1], f'{model:.1f}', f'{year} transport capital costs')
            self.assertEqual(result[year]['O&M costs'][-1], 'in capital*', f'{year} transport O&M costs')
            self.assertTrue(any('investment and maintenance' in note for note in result[year + ' notes']), f'{year}: explanation missing')

    def test_sankey_shows_all_primary_energy(self):
        result = self.get('plots')
        for year in ('BY', 'SC'):
            self.assertAlmostEqual(result['sankey ' + year]['shown'], result['sankey ' + year]['model'], places=6, msg=year)

    def test_duration_curve_axis_shows_the_whole_curve(self):
        result = self.get('plots')['duration curve']
        self.assertLessEqual(result['axis'][0], result['curve'][0])
        self.assertGreaterEqual(result['axis'][1], result['curve'][1])

    def test_cut_off_resource_bars_are_labelled(self):
        result = self.get('plots')['resources']
        self.assertTrue(result['longest bars'], 'the default data has resources above 500 PJ')
        for value in result['longest bars']:
            self.assertTrue(any(f'{value:.0f}' in label for label in result['labels']), f'no label for the {value:.0f} PJ bar')

    def test_invalid_input_is_refused_with_an_explanation(self):
        result = self.get('input_validation')
        for name in ('share -5', 'share 120'):
            self.assertIn('between 0 % and 100 %', result[name]['message'] or '', name)
            self.assertEqual(result[name]['shown'], result['original share'], f'{name} must not change the table')
        self.assertIn('number', result['share abc']['message'] or '')
        self.assertIn('utilisation', (result['utilisation 0']['message'] or '').lower())
        self.assertIn('0 €/GJ or more', result['price -3']['message'] or '')
        self.assertIn('CO2 price', result['co2 price -10']['message'] or '')
        self.assertEqual(result['co2 price -10']['co2 price now'], 50)

    def test_valid_input_is_accepted(self):
        result = self.get('input_validation')
        self.assertEqual((result['share 23,5']['message'], result['share 23,5']['shown']), (None, '23.50%'))
        self.assertEqual((result['share 7.25%']['message'], result['share 7.25%']['shown']), (None, '7.25%'))
        self.assertEqual((result['price 12']['message'], result['price 12']['shown']), (None, '12.00'))

    def test_update_is_refused_when_shares_do_not_add_up(self):
        result = self.get('input_validation')['update with wrong sums']
        errors = [text for kind, title, text in result['messages'] if kind == 'showerror']
        self.assertEqual(len(errors), 1, result['messages'])
        for expected in ('Residential', 'Car', '100 %'):
            self.assertIn(expected, errors[0])
        self.assertTrue(result['file unchanged'], 'a refused update must not save the scenario')

    def test_electricity_shares_above_100_percent_are_shown_as_export(self):
        labels = self.get('input_validation')['electricity labels']
        self.assertTrue(any('export' in label.lower() for label in labels), labels)

    def test_production_amounts_are_shown_next_to_the_shares(self):
        result = self.get('input_validation')
        for name in ('electricity', 'district heat'):
            amounts = result[name + ' amounts']
            self.assertEqual(amounts['shown'], amounts['model'], name)
            self.assertEqual(len(amounts['shown']), amounts['rows in shares table'], f'{name}: one amount per row')

    def test_transport_fuel_sums_are_shown_next_to_the_shares(self):
        # the check has added 2 % electricity to cars, so the car's sum is 102 % and shown in red
        sums = self.get('input_validation')['transport fuel sums']
        self.assertEqual(len(sums), 13, 'a sum for each of the 10 transport types and an empty cell for the 3 group rows')
        self.assertEqual([tags for value, tags in sums if value == ''], [['group']] * 3)
        self.assertEqual(sums[1], ['102.00%', ['error']])
        self.assertEqual([value for value, tags in sums if tags == []], ['100.00%'] * 9)

    def test_save_scenario_as(self):
        self.assertEqual(self.get('scenario_files')['saved shares'], [0.146, 0.334])

    def test_reset_to_default_scenario(self):
        result, expected = self.get('scenario_files')['after reset'], helpers.load_expected('default')
        for year in ('BY', 'SC'):
            self.assertEqual(helpers.compare(expected[year], result[year], rel_tol=1e-8), [], year)

    def test_open_scenario_gives_same_results_as_fresh_start(self):
        result = self.get('scenario_files')
        self.assertEqual(result['year shown'], '2035')
        for year in ('BY', 'SC'):
            self.assertEqual(helpers.compare(result['fresh start after open'][year], result['after open'][year], rel_tol=1e-8), [], year)

    def test_opening_a_file_that_is_not_a_scenario_is_refused(self):
        result = self.get('scenario_files')
        errors = [text for kind, title, text in result['bad file messages'] if kind == 'showerror']
        self.assertEqual(len(errors), 1, result['bad file messages'])
        self.assertIn('not a pySTREAM scenario', errors[0])
        self.assertEqual(helpers.compare(result['after open'], result['after bad file']), [], 'nothing may change')

    def test_time_series_is_calculated_when_its_page_is_opened(self):
        result = self.get('time_series_on_demand')
        self.assertFalse(result['calculated at startup'], 'the time series should not slow down the start')
        self.assertTrue(result['calculated when page opened'])
        self.assertFalse(result['kept after update'], 'after Update Scenario the old time series must not be shown')

    def test_start_message_is_removed_when_started(self):
        self.assertFalse(self.get('time_series_on_demand')['start message still shown'])

    def test_window_can_be_scrolled_on_a_small_screen(self):
        result = self.get('small_screen')
        self.assertTrue(result['horizontal scrollbar shown'])
        self.assertTrue(result['vertical scrollbar shown'])
        self.assertEqual(result['visible part at start'][0], 0)
        self.assertGreater(result['visible part after mouse wheel'][0], 0, 'the mouse wheel should scroll down')
        self.assertEqual(result['visible part at the bottom'][1], 1, 'the bottom of the layout should be reachable')

    def test_layout_is_scaled_to_fit_the_screen(self):
        # (The window cannot be made larger than the real screen the tests run on, so 'no scrollbars' is only
        # checked when the window got room for the whole layout.)
        for check in ('screen_laptop', 'screen_design', 'screen_large'):
            with self.subTest(check):
                result = self.get(check)
                width, height = result['screen']
                scale = min(width / 1920, height / 1080)
                self.assertAlmostEqual(result['scale'], scale)
                self.assertEqual(result['layout'], [int(1920 * scale), int(1080 * scale)])
                self.assertAlmostEqual(result['tk scaling'], scale * 96 / 72, places=2, msg='text should be scaled like the layout')
                self.assertEqual(result['table row height'], round(23 * scale))
                self.assertEqual(result['cut off text'], [])
                window_width, window_height = result['window inside']
                if window_width >= result['layout'][0]:
                    self.assertFalse(result['horizontal scrollbar shown'], 'the whole layout should fit the screen')
                if window_height >= result['layout'][1]:
                    self.assertFalse(result['vertical scrollbar shown'], 'the whole layout should fit the screen')

    def test_text_stays_readable_on_a_very_small_screen(self):
        result = self.get('screen_tiny')
        self.assertEqual(result['scale'], 0.6, 'the layout should not get smaller than the minimum')
        self.assertTrue(result['horizontal scrollbar shown'])
        self.assertTrue(result['vertical scrollbar shown'])
        self.assertEqual(result['cut off text'], [])

    def check_friendly_error(self, check, expected_text):
        result = self.get(check)
        errors = [text for kind, title, text in result['messages'] if kind == 'showerror']
        self.assertEqual(len(errors), 1, f'expected one error message, got {result["messages"]}')
        self.assertIn(expected_text, errors[0])
        self.assertIn('error log', errors[0], 'the message should say where the details are')
        self.assertIn('Traceback', result['log'] or '', 'the details should be in the error log')

    def test_friendly_message_when_data_file_is_open_in_excel(self):
        self.check_friendly_error('file_open_in_excel', 'open in Excel')

    def test_friendly_message_when_a_data_file_is_missing(self):
        self.check_friendly_error('missing_data_file', 'Assumptions.xlsx')

    def check_sankey_fallback(self, check):
        texts = self.get(check)['button_texts']
        browser_buttons = [t for t in texts if 'browser' in t]
        self.assertEqual(len(browser_buttons), 2, f'expected a browser button for both Sankey diagrams, got {texts}')

    def test_sankey_falls_back_to_browser_when_image_fails(self):
        # e.g. Google Chrome, which Kaleido needs, is not installed
        self.check_sankey_fallback('sankey_image_fails')

    def test_sankey_falls_back_to_browser_when_image_hangs(self):
        self.check_sankey_fallback('sankey_image_hangs')

    def test_closing_the_window_stops_pystream(self):
        self.assertTrue(self.get('close_window')['stopped_when_closed'])

    def test_update_time_series_button_works(self):
        self.assertIsNone(self.get('update_time_series_button')['exception'])

    def test_update_scenario_year(self):
        result = self.get('update_scenario_year')
        self.assertEqual(result['exceptions'], [])
        self.assertEqual(len(result['invalid_input']['messages']), 3, 'each invalid year should give a message')
        self.assertTrue(result['invalid_input']['file_unchanged'], 'an invalid year must not change the data file')
        differences = helpers.compare(result['fresh_start'], result['after_update'], rel_tol=1e-8)
        if differences:
            self.fail(f'{len(differences)} results differ after changing the year, e.g.\n  ' + '\n  '.join(differences[:15]))

    def test_update_scenario_gives_same_results_as_fresh_start(self):
        # After 'Update Scenario' every result must be what the model gives when
        # started fresh with the saved data; anything else is a stale result.
        result = self.get('update_matches_fresh_start')
        self.assertEqual(result['messages'], [], 'this change should not cause any message')
        differences = helpers.compare(result['fresh_start'], result['after_update'], rel_tol=1e-8)
        if differences:
            self.fail(f'{len(differences)} results differ after Update Scenario, e.g.\n  ' + '\n  '.join(differences[:15]))


if __name__ == '__main__':
    unittest.main()
