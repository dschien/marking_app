from collections import defaultdict

import xlrd
import numpy as np

__author__ = 'schien'

sections = {23: "Uncertainty", 38: "Improvements", 46: "Interpretation", 56: "Tool", 65: "Writing style"}
sortedKeys = sorted(sections.keys())

CAT = ''
GROUP = 'group'
MASTER = 'master'
MARK = 'Mark'
STUDENT = 'student'
EMAIL = 'email'

TABLE_STRUCT = {
    CAT: 0,
    MARK: 1,
    GROUP: 1,
    MASTER: 2,
    STUDENT: 0,
    EMAIL: 10
}

import logging


class MarkingData(object):
    def __init__(self, file='marking_table.xlsx'):
        self.wb = self.load_workbook(file)

        self.master_students = self.load_master_students()
        self.bachelor_students = self.load_bachelor_students()

        self.groups = self.load_group_names()
        self.masters_groups, self.bachelors_groups = self.load_group_membership()

        self.marks = self.load_marks()
        self.master_mark_weights, self.bachelors_mark_weights = self.load_mark_weigths()

        self.masters_max_mark, self.bachelor_max_mark = self.load_achievable_marks()

        self.final_marks = self.calculate_marks()


    def load_group_membership(self):
        master_groups = []
        bachelor_groups = []
        for group, members in self.groups.items():

            # if there is at least one MSc student, they all are marked like MSc students
            if any([member for member in members if member in self.master_students]):
                master_groups.append(group)
            else:
                # if there are no MSc students, then there should be at least one BA student
                if any([member for member in members if member in self.bachelor_students]):
                    bachelor_groups.append(group)
                else:
                    # or we did not identify any group members
                    logging.warn("Group could not be allocated to BA/MSc {}".format(group))
        return master_groups, bachelor_groups

    def load_master_students(self):
        """

        :return:
        """
        master_students = set()
        sheet = self.wb.sheet_by_name("MSc")

        for row in sheet.col(0):
            master_students.add(row.value)
        return master_students

    def load_bachelor_students(self):
        """

        :return:
        """
        sheet = self.wb.sheet_by_name("BA")
        bachelor_students = set()
        for row in sheet.col(0):
            bachelor_students.add(row.value)

        return bachelor_students

    def get_groups(self):
        # names = [name if type(name[0]) is not float else (str(int(name[0])), name[1]) for name in names]
        # names = [name for name in names if len(name) > 0]
        return self.groups.keys()

    def load_workbook(self, file):
        wb = xlrd.open_workbook(file)
        return wb

    def load_group_names(self, start_rowx=0, end_rowx=None):
        sheet_names = self.wb.sheet_names()
        groups = defaultdict(list)
        for name in sheet_names:
            try:
                n = int(name)
                # groups.append(n)
                sheet = self.wb.sheet_by_name(name)
                for candidate_row in range(0, 6):
                    if sheet.cell_type(candidate_row, 0) in (xlrd.XL_CELL_NUMBER, ):
                        groups[n].append(sheet.cell(candidate_row, 0).value)
                if not n in groups:
                    logging.warn("No candidates found in group {}".format(name))
            except ValueError as e:
                # print e
                pass
        return groups

    def get_group_for_candidate(self, candidate):
        for group, candidates in self.groups.iteritems():
            if candidate in candidates:
                return group

    def load_marks(self):
        marks = defaultdict(dict)
        for group in self.groups:
            sh = self.wb.sheet_by_name(str(group))

            col_idx = 1

            for row_idx in [23, 38, 46, 56, 65]:
                cell = sh.cell(row_idx, col_idx)
                current_section = sections[row_idx]
                if sh.cell_type(row_idx, col_idx) == xlrd.XL_CELL_NUMBER:
                    marks[group][current_section] = int(cell.value)
                else:
                    if not (row_idx == 38 and group in self.bachelors_groups):
                        logging.warning('No mark for section {} of group {}'.format(current_section, group))

        return marks


    def get_marks_as_array(self, group):
        sortedSections = [sections.get(k) for k in sortedKeys]
        if group in self.bachelors_groups:
            sortedSections.remove(sections[38])
        group_marks = [self.marks[group].get(k) for k in sortedSections]
        return group_marks

    def get_weighted_marks_as_array(self, group):
        marks = self.get_marks_as_array(group)
        if group in self.bachelors_groups:
            return np.asarray(marks) * 10 / self.bachelors_mark_weights
        else:
            return np.asarray(marks) * 10 / self.master_mark_weights

    def load_achievable_marks(self):

        sheet = self.wb.sheet_by_name("marking")
        master_row_column = 2
        bachelor_mark_column = 1
        masters_max = sheet.cell(6, master_row_column).value
        ba_max = sheet.cell(6, bachelor_mark_column).value

        return masters_max, ba_max

    def load_mark_weigths(self):
        sheet = self.wb.sheet_by_name("marking")
        master_marks = []
        bachelors_marks = []
        master_row_column = 2
        for row in range(1, 6):
            master_marks.append(sheet.cell(row, master_row_column).value)
        # skip improvements
        bachelor_mark_column = 1
        bachelors_marks.append(sheet.cell(1, bachelor_mark_column).value)
        for row in range(3, 6):
            bachelors_marks.append(sheet.cell(row, bachelor_mark_column).value)

        return master_marks, bachelors_marks

    def calculate_marks(self):
        final_marks = {}

        for group, candidates in self.groups.items():
            marks = self.get_marks_as_array(group)
            for candidate in candidates:
                if group in self.masters_groups:
                    final_marks[candidate] = sum(marks) / self.masters_max_mark * 100

                else:
                    final_marks[candidate] = sum(marks) / self.bachelor_max_mark * 100

        return final_marks