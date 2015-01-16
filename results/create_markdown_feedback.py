import csv
import logging
from os import path
import sys

from create_marks_charts import Analysis


__author__ = 'schien'

import xlrd

wb = xlrd.open_workbook('../marking_table.xlsx')
sheet_names = wb.sheet_names()
sections = {23: "Uncertainty", 28: "Improvements", 46: "Interpretation", 56: "Tool", 65: "Writing style"}
analysis = Analysis(path.join(path.dirname(path.dirname(__file__)), 'marking_table.xlsx'))

feedback_dir = path.join(path.dirname(path.dirname(__file__)), 'feedback')


def create_markdown(candidate, group):
    try:
        output = open(path.join(feedback_dir, '{:.0f}.md'.format(candidate)), 'w')

        output.write("# Assignment Feedback \n")

        sheet = wb.sheet_by_name(str(group))
        # output = StringIO.StringIO()
        output.write("## Candidate  {:.0f}\n".format(candidate))

        for i in range(17, 72):

            row_idx = i
            col_idx = 1
            cell = sheet.cell(row_idx, col_idx)
            if row_idx in [23, 28, 46, 56, 65]:
                output.write("\n## {}\n".format(sections[row_idx]))
                # output.write ('-'*40 )
                output.write("\n")
            if not sheet.cell_type(row_idx, col_idx) in (
                    xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_NUMBER):
                output.write('%s\n' % cell.value)

        output.write('\n## %s\n' % 'Results')

        output.write('\n### Final Mark: {:.0f}\n'.format(analysis.data.final_marks[candidate]))


        output.write(
            '\n![Comparison of performance per marking category with course average](images/{:.0f}.png)\n'.format(
                candidate))

        # output.write('\n### %s\n' % 'Weights')
        output.write('\n![Weightings of marks for bachelor level groups](../weights_ba.png)\n')
        output.write('\n![Weightings of marks for M-level groups](../weights_m.png)\n')

        output.close()
    except ValueError as e:
        # print e
        logging.warn("Problem with candiate {}: {}".format(candidate,e))


def process_candidates():

    covered_candidates = []
    missing_candidates = []



    for candidate in analysis.data.bachelor_students.union(analysis.data.master_students):

        group = analysis.data.get_group_for_candidate(candidate)
        if not group:
            logging.warn("No group found for candidate {}".format(candidate))
            missing_candidates.append(candidate)
            continue

        # if group in analysis.data.masters_groups:
        # in_master_group.add(candidate)
        # continue
        create_markdown(candidate, group)
        covered_candidates.append(("{:.0f}".format(candidate), "{:.0f}".format(analysis.data.final_marks[candidate]),))

        # for candidate in analysis.data.master_students.union(in_master_group):
    covered_candidates.sort()
    # pprint(groups,width=2)

    writer = csv.writer(sys.stdout)

    print "Covered candidates"
    for item in covered_candidates:
        writer.writerow(item)

    print "missing_candidates"
    for item in missing_candidates:
        writer.writerow(["{:.0f}".format(item)])



if __name__ == '__main__':
    process_candidates()