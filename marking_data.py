import itertools

__author__ = 'schien'
import os
import collections

import xlrd
from fuzzywuzzy import process


basedir = os.path.abspath(os.path.dirname(__file__))


class MarkingData(object):
    def __init__(self, app):
        wb = xlrd.open_workbook(os.path.join(basedir, 'marking_table.xlsx'))
        sheet_names = wb.sheet_names()
        self.app = app
        # dict of groups with dict of sections with list of entries
        self.groups = {}
        # dict of groups with dict of sections with mark
        self.marks = {}

        sections = {23: "Uncertainty", 38: "Improvements", 46: "Interpretation", 56: "Tool", 65: "Writing style"}
        for name in sheet_names:

            try:
                n = int(name)
                self.groups[n] = {}
                self.marks[n] = {}

                sheet = wb.sheet_by_name(name)

                current_section = None
                for i in range(23, 72):

                    row_idx = i
                    col_idx = 1
                    cell = sheet.cell(row_idx, col_idx)
                    if row_idx in [23, 38, 46, 56, 65]:
                        if current_section:
                            self.groups[n][current_section] = current_section_data
                        current_section = sections[row_idx]
                        current_section_data = []
                        if sheet.cell_type(row_idx, col_idx) == xlrd.XL_CELL_NUMBER:
                            self.marks[n][current_section] = int(cell.value)
                        else:
                            app.logger.warning('No mark for section {} of group {}'.format(current_section, n))

                    if not sheet.cell_type(row_idx, col_idx) in (
                            xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_NUMBER):
                        current_section_data.append(cell.value)

                self.groups[n][current_section] = current_section_data
            except ValueError as e:
                # print e
                pass


    def search_token(self, token, section, limit=10):
        sections = ["Uncertainty", "Improvements", "Interpretation", "Tool", "Writing style"]

        # flatten array

        choices = {}
        c = collections.Counter()
        for group, group_data in self.groups.items():
            for entry in group_data[section]:
                c.update(str(group))

                choices[str(group) + "," + str(c[str(group)])] = entry
        print choices
        results = process.extract(token, choices, limit=limit)

        groups = {}
        uniquekeys = []

        data = sorted(results, key=lambda item: item[2].split(',')[0])
        for key, grp in itertools.groupby(data, key=lambda item: item[2].split(',')[0]):
            uniquekeys.append(key)
            groups[key.split(',')[0]] = list(grp)

        res = []
        for group, vals in groups.items():
            max_score = 0

            for val in vals:
                if val[1] > max_score:
                    max_score = val[1]

            res.append((group, vals, max_score))

        res = sorted(res, key=lambda item: item[2], reverse=True)
        return res

    def filter(self, mark, section):
        results = {}
        marks = {}
        for group, sections in self.marks.items():
            if not section in sections:
                self.app.logger.warning('No mark for section {} of group {}'.format(section, group))
            if section in sections and sections[section] >= mark:
                results[group] = self.groups[group][section]
                marks[group] = sections[section]

        return results, marks