import csv
import sys

import xlrd


wb = xlrd.open_workbook('marking_table.xlsx')
sheet_names = wb.sheet_names()
# print('Sheet Names', sheet_names)
groups = []

sections = {23: "Uncertainty", 28: "Improvements", 46: "Interpretation", 56: "Tool", 65: "Writing style"}
for name in sheet_names:

    try:
        n = int(name)
        groups.append(n)

        output = open('response/%s.md' % n, 'w')
        sheet = wb.sheet_by_name(name)
        # output = StringIO.StringIO()
        for candidate_row in range(0, 6):

            if not sheet.cell_type(candidate_row, 0) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                output.write("{}\n".format(str(sheet.cell(candidate_row, 0).value)))

        for i in range(17, 72):


            row_idx = i
            col_idx = 1
            cell = sheet.cell(row_idx, col_idx)
            if row_idx in [23, 28, 46, 56, 65]:
                output.write("\n# {}\n".format(sections[row_idx]))
                # output.write ('-'*40 )
                output.write("\n")
            if not sheet.cell_type(row_idx, col_idx) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_NUMBER):
                output.write('%s\n' % cell.value)

        output.write('\n# %s\n' % 'Results')
        output.write('\n![Radar](response/images/%s.png)\n' % name)

        output.close()
    except ValueError as e:
        # print e
        pass

groups.sort()
# pprint(groups,width=2)

writer = csv.writer(sys.stdout)

for item in groups:
    writer.writerow([item])


