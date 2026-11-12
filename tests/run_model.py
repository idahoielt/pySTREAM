"""Run pySTREAM's model in a copy of the program and write its results to a JSON file.

Used by the tests (see helpers.run_model); runs in its own process because the
model calculates everything when Backend.py is imported.

Usage: python run_model.py <program_dir> <output.json> [--time-series] [--cwd <folder>]
"""
import argparse
import json
import os
import sys
import warnings

import helpers

parser = argparse.ArgumentParser()
parser.add_argument('program_dir')
parser.add_argument('output_file')
parser.add_argument('--time-series', action='store_true')
parser.add_argument('--cwd', help='folder to start the model from (default: the program folder)')
args = parser.parse_args()

warnings.simplefilter('ignore')
os.chdir(args.cwd or args.program_dir)
sys.path.insert(0, os.path.join(args.program_dir, 'program'))

import Backend

results = helpers.collect_outputs(Backend)
if args.time_series:
    import Time_Series
    B = Backend
    df_hour, technologies_used = Time_Series.run_timeSeries(B.df_Power_SC, B.df_FuelPrice_SC, B.df_FEC_SC, B.total_electricity_demand_SC,
                                                            B.el_transport_SC, B.el_DH_SC, B.el_FR_SC)
    results['time_series'] = helpers.collect_time_series(df_hour, technologies_used)

with open(args.output_file, 'w', encoding='utf-8') as f:
    json.dump(results, f)
