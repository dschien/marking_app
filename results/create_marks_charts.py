from __future__ import division
import logging
from os import path

from load_excel_data import MarkingData, sections, sortedKeys
from plot_radar_results import plot_radar

images_dir = path.join(path.dirname(path.dirname(__file__)), 'feedback/images')

__author__ = 'schien'
import numpy as np


class Analysis(object):
    def __init__(self, file):
        self.data = MarkingData(file)
        # self.names = self.data.get_groups()

        self.categories_m = [sections.get(k) for k in sortedKeys]
        self.categories_b = [sections.get(k) for k in sortedKeys]
        self.categories_b.remove(sections[38])

        self.collect_average_values()

        self.average_m, self.average_b = self.collect_average_values()

    def collect_average_values(self):
        self.average_m = []
        self.average_b = []

        master_scores = []
        bachelor_scores = []

        for group in self.data.masters_groups:
            master_scores.append(self.data.get_weighted_marks_as_array(group))
        for group in self.data.bachelors_groups:
            bachelor_scores.append(self.data.get_weighted_marks_as_array(group))

        average_m = np.average(np.asarray(master_scores), axis=0)
        average_b = np.average(np.asarray(bachelor_scores), axis=0)
        return average_m, average_b


    def plot_groups(self):
        for group in self.data.masters_groups:
            scores = self.data.get_weighted_marks_as_array(group)
            plot_radar(self.categories_m, self.average_m, scores, path.join(images_dir, str(group) + ".png"), group)
        for group in self.data.bachelors_groups:
            scores = self.data.get_weighted_marks_as_array(group)
            plot_radar(self.categories_b, self.average_b, scores, path.join(images_dir, str(group) + ".png"), group)

    def plot_master_candidates(self, in_master_group):
        for candidate in self.data.master_students.union(in_master_group):
            # print candidate
            group = self.data.get_group_for_candidate(candidate)
            if not group:
                logging.warn("No group found for MSc candidate {:.0f}".format(candidate))
                continue
            # print group
            scores = self.data.get_weighted_marks_as_array(group)
            plot_radar(self.categories_m, self.average_m, scores, path.join(images_dir, "{:.0f}".format(candidate)  + ".png"), group)

    def plot_bachelor_candidates(self):
        in_master_group = set()

        for candidate in self.data.bachelor_students:

            group = self.data.get_group_for_candidate(candidate)
            if not group:
                logging.warn("No group found for BA candidate {:.0f}".format(candidate))
                continue

            if group in self.data.masters_groups:
                in_master_group.add(candidate)
                continue
            scores = self.data.get_weighted_marks_as_array(group)
            plot_radar(self.categories_b, self.average_b, scores,
                       path.join(images_dir, "{:.0f}".format(candidate) + ".png"), group)

        return in_master_group


if __name__ == '__main__':
    analysis = Analysis(path.join(path.dirname(path.dirname(__file__)), 'marking_table.xlsx'))
    in_master_group = analysis.plot_bachelor_candidates()
    analysis.plot_master_candidates(in_master_group)
    # analysis.plot_groups()