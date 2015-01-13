from os import path

import numpy as np

from results import test_dir
from results.create_marks_charts import Analysis
from results.load_excel_data import MarkingData

__author__ = 'schien'

import unittest


class MyTestCase(unittest.TestCase):
    
    def test_marking_data_init(self):
        data = MarkingData(path.join(test_dir, 'test_marking_data.xlsx'))

        assert set(data.masters_groups) == {41, 42}
        assert set(data.bachelors_groups) == {44}
        assert set(data.groups.keys()) == {41, 42, 44}

        assert data.get_marks_as_array(41) == [4, 0, 6, 1, 3]

        assert all(
            [abs(i - j) < 0.0001 for i, j in zip(data.get_weighted_marks_as_array(41), [3.636363636, 0, 6, 2, 10])])

        assert data.get_marks_as_array(42) == [11, 11, 10, 5, 3]
        assert all(data.get_weighted_marks_as_array(42) == [10, 10, 10, 10, 10])
        assert all(data.get_weighted_marks_as_array(44) == [10, 6, 0, 10])



    def test_chart_init(self):
        analysis = Analysis(path.join(test_dir, 'test_marking_data.xlsx'))

        assert all([abs(i - j) < 0.0001 for i, j in
                    zip(analysis.average_m, np.average([[10, 10, 10, 10, 10], [3.636363636, 0, 6, 2, 10]], axis=0))])

        assert (analysis.average_b == [10, 6, 0, 10]).all()

    def test_plotting(self):
        analysis = Analysis(path.join(test_dir, 'test_marking_data.xlsx'))
        analysis.plot_groups()


if __name__ == '__main__':
    unittest.main()
