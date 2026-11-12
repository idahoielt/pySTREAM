"""Re-run all test scenarios and store their results as the new expected values.

Only run this when a change is *meant* to change the model results. It prints the
largest changes so they can be documented (commit message and docs/CHANGELOG.md).

Usage (from the project folder): python tests/update_expected.py
"""
import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers


def flatten(results, path=''):
    if isinstance(results, dict):
        for key, value in results.items():
            yield from flatten(value, f'{path}/{key}')
    else:
        yield path, results


def main():
    for scenario, results in helpers.run_all_scenarios().items():
        if not os.path.exists(helpers.expected_file(scenario)):
            helpers.save_expected(scenario, results)
            print(f'{scenario}: created')
            continue
        old = dict(flatten(helpers.load_expected(scenario)))
        changes = []
        for path, new_value in flatten(results):
            old_value = old.get(path)
            if isinstance(old_value, float) and isinstance(new_value, float):
                if not math.isclose(old_value, new_value, rel_tol=1e-9, abs_tol=1e-9):
                    changes.append((abs(new_value - old_value), path, old_value, new_value))
            elif old_value != new_value:
                changes.append((math.inf, path, old_value, new_value))
        helpers.save_expected(scenario, results)
        print(f'{scenario}: {len(changes)} values changed')
        for _, path, old_value, new_value in sorted(changes, key=lambda c: -c[0])[:15]:
            print(f'    {path}: {old_value!r} -> {new_value!r}')


if __name__ == '__main__':
    main()
