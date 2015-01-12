import csv
import sys

import xlrd
from collections import defaultdict


wb = xlrd.open_workbook('marking_table.xlsx')
sheet_names = wb.sheet_names()
# print('Sheet Names', sheet_names)

groups = defaultdict(list)

for name in sheet_names:

    try:
        n = int(name)

        sheet = wb.sheet_by_name(name)
        # output = StringIO.StringIO()
        for candidate_row in range(1, 3):
            if sheet.cell_type(candidate_row, 0) in (xlrd.XL_CELL_NUMBER, ):
                groups[name].append(sheet.cell(candidate_row, 0).value)

    except ValueError as e:
        # print e
        pass


writer = csv.writer(sys.stdout)
print groups
for group, candidates in groups.items():
    writer.writerow([group] + candidates)


