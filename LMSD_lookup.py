"""
LMSD database lookup script

MIT License

Copyright (c) 2017 Thomas Jungers

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import sys
from os.path import basename
from openpyxl import load_workbook
import collections
from tqdm import tqdm
import re
import sqlite3
import html_msg


Adduct = collections.namedtuple('Adduct', ['mass', 'charge'])
ADDUCTS = {
    'H': Adduct(1.007276, 1),
    'Na': Adduct(22.989218, 1),
    'K': Adduct(38.963158, 1),
    '2H': Adduct(2 * 1.007276, 2)
}
ADDUCT_NAMES = " - ".join(sorted(ADDUCTS, key=ADDUCTS.get))

INTENSITY_THRESHOLD = 100000

NUM_PAD = {  # type of number: (non-decimal digits, decimal digits)
    'mass': (5, 5),
    'intensity': (9, 1),
    'results': (2, 0),
    'error_abs': (1, 5),
    'error_rel': (4, 4)
}


deutRe = re.compile("([0-9A-Za-z]|^)D([1-9A-Z]|$)")
chemRe = re.compile("(?<=[A-Za-z])([0-9]+)")


def noDeuterium(item):
    return deutRe.search(item) is None


def chem_fmt(s):
    return chemRe.sub(r'<sub>\1</sub>', s)


def num_fmt(num, n_type):
    fmt_total = NUM_PAD[n_type][0] + NUM_PAD[n_type][1] + 1
    fmt_dec = NUM_PAD[n_type][1]
    fmt = " {}.{}".format(fmt_total, fmt_dec)
    return ("{:{}f}".format(num, fmt)
                    .replace(" ", "&#x2007;"))


def isascii(s):
    return len(s) == len(s.encode())


def get_input(prompt, type_=None, min_=None,
              max_=None, range_=None, tests=None):
    if min_ is not None and max_ is not None and max_ < min_:
        raise ValueError("min_ must be less than or equal to max_.")
    while True:
        ui = input(prompt)
        if type_ is not None:
            try:
                ui = type_(ui)
            except ValueError:
                print("Input type must be {0}.".format(type_.__name__))
                continue
        if tests is not None:
            test_failed = False
            for test in tests:
                if not test(ui):
                    print("Input must satisfy '{}'".format(test.__name__))
                    test_failed = True
                    break
            if test_failed:
                continue
        if max_ is not None and ui > max_:
            print("Input must be less than or equal to {0}.".format(max_))
        elif min_ is not None and ui < min_:
            print("Input must be greater than or equal to {0}.".format(min_))
        elif range_ is not None and ui not in range_:
            if isinstance(range_, range):
                template = "Input must be between {0.start} and {0.stop}."
                print(template.format(range_))
            else:
                template = "Input must be {0}."
                if len(range_) == 1:
                    print(template.format(*range_))
                else:
                    print(template.format(" or ".join(
                        (", ".join(map(str, range_[:-1])), str(range_[-1])))
                    ))
        else:
            return ui


try:
    wb = load_workbook(sys.argv[1])
except Exception as exc:
    print("Could not load Excel file: {}".format(exc))
    exit()

# trim the path to output the report to the script directory
# input_fname = basename(sys.argv[1])
# keep the path to output the report to the data file directory
input_fname = sys.argv[1]

# db in a fixed directory (when launched from .bat)
conn = sqlite3.connect(r'C:\LMSD_lookup\LMSD.db')
# db in current directory (when launched from console)
# conn = sqlite3.connect(r'LMSD.db')
conn.row_factory = sqlite3.Row
conn.create_function("NODEUTERIUM", 1, noDeuterium)
db = conn.cursor()

sheet_names = wb.get_sheet_names()

while True:
    sheet_num = None

    for i, sheet in enumerate(sheet_names):
        print("{:2d}. {}".format(i, sheet))
    print(" Q. Exit")
    sheet_num = input("Sheet number: ")

    if sheet_num.lower() == 'q':
        exit()

    try:
        sheet_name = sheet_names[int(sheet_num)]
        sheet = wb[sheet_name]
    except Exception:
        continue

    print("{} selected".format(sheet))

    mass_col = get_input("Mass column: ", tests=[isascii, str.isalpha]).upper()
    int_col = get_input("Intensity column: ",
                        tests=[isascii, str.isalpha]).upper()
    start_row = get_input("Start row: ", type_=int, min_=1)
    end_row = get_input("End row: ", type_=int, min_=1)

    relative_error = input("Relative (R) or absolute (A) error? ")
    relative_error = relative_error.upper() == 'R'
    if relative_error:
        mass_error = get_input("Relative error (ppm): ", type_=float, min_=0)
    else:
        mass_error = get_input("Absolute error (amu): ", type_=float, min_=0)

    if start_row > end_row:
        start_row, end_row = end_row, start_row

    row_range = range(start_row, end_row + 1)
    str_range = [
        "{}{}".format(mass_col, start_row),
        "{}{}".format(mass_col, end_row)
    ]

    COUNT = {
        'total': 0,
        'masses': 0,
        'matches': 0
    }

    report_fname = ("{file}__{sheet}__{range_start}-{range_end}"
                    "__{error}-{rel}.html").format(
        file=input_fname, sheet=sheet_name,
        range_start=str_range[0], range_end=str_range[1],
        error=mass_error, rel='r' if relative_error else 'a'
    )
    with open(report_fname, 'w') as f:
        f.write(html_msg.doctype.format(
            file=input_fname, sheet=sheet_name,
            range_start=str_range[0],
            range_end=str_range[1],
            adducts=ADDUCT_NAMES,
            error=mass_error,
            error_unit=('ppm' if relative_error else 'uma'),
            threshold=INTENSITY_THRESHOLD)
        )
        f.write(html_msg.mass_start)

        for row in tqdm(row_range):
            try:
                mass = float(sheet["{}{}".format(mass_col, row)].value)
                intensity = float(sheet["{}{}".format(int_col, row)].value)
            except Exception:
                print(("\rIncorrect values in row {}; "
                       "skipping it\033[K").format(row))
                continue

            COUNT['masses'] += 1

            isomers = collections.defaultdict(list)
            count = 0

            for adduct_name, adduct in ADDUCTS.items():
                cor_mass = mass * adduct.charge - adduct.mass
                if relative_error:
                    low_mass = cor_mass / (1 + mass_error * 1e-6)
                    high_mass = cor_mass / (1 - mass_error * 1e-6)
                else:
                    low_mass = cor_mass - mass_error
                    high_mass = cor_mass + mass_error

                db_res = db.execute(
                    ('SELECT *, ? as adduct, ? as cor_mass FROM LMSD '
                     'WHERE MASS BETWEEN ? AND ? '
                     'AND NODEUTERIUM(FORMULA) ORDER BY MASS'),
                    (adduct_name, cor_mass, low_mass, high_mass)
                )

                for db_row in db_res:
                    count += 1
                    isomers[db_row['FORMULA']].append(db_row)

            if isomers:
                COUNT['matches'] += 1
                COUNT['total'] += count
                bold = "bold" if intensity >= INTENSITY_THRESHOLD else "gray"
                f.write(html_msg.mass.format(
                    mass=num_fmt(mass, 'mass'),
                    intensity=num_fmt(intensity, 'intensity'),
                    total=num_fmt(count, 'results'),
                    isomers=len(isomers),
                    bold=bold)
                )
                f.write(html_msg.formula_start)

                for formula in sorted(
                    isomers,
                    key=lambda x: isomers[x][0]['MASS']
                ):
                    group = isomers[formula]
                    error = (float(group[0]['MASS']) -
                             float(group[0]['cor_mass']))
                    error_rel = error / float(group[0]['MASS'])
                    f.write(html_msg.formula.format(
                        formula=chem_fmt(formula),
                        mass=num_fmt(float(group[0]['MASS']), 'mass'),
                        adduct=group[0]['adduct'],
                        error_abs=num_fmt(error, 'error_abs'),
                        error_rel=num_fmt(error_rel * 1e6, 'error_rel'),
                        count=len(group))
                    )
                    f.write(html_msg.entry_start)

                    for isomer in group:
                        f.write(html_msg.entry.format(**isomer))

                    f.write(html_msg.entry_end)

                f.write(html_msg.formula_end)

        f.write(html_msg.mass_end)
        f.write(html_msg.end.format(**COUNT))

        print("Report written to {}".format(report_fname))

conn.close()
