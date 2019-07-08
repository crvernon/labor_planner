"""test_builder.py

Tests for BuildStaffWorkbooks class.

@author Chris R. Vernon (chris.vernon@pnnl.gov)
@license BSD 2-Clause

"""

import os
import unittest

from labor_planner.config_reader import ReadConfig
from labor_planner.labor_builder.build_staff_workbooks import BuildStaffWorkbooks


class TestBuilder(unittest.TestCase):
    """Test BuildStaffWorkbooks attributes."""

    TEST_DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
    TEST_CONFIG_FILE = os.path.join(TEST_DATA_DIR, 'config_build.yml')
    TEST_CONFIG_OBJ = ReadConfig(TEST_CONFIG_FILE)
    TEST_READ_OBJ = BuildStaffWorkbooks(TEST_CONFIG_OBJ)

    def test_num_months_wkg(self):
        """Check the number of months in the working hours file derived list."""

        n_months = len(TestBuilder.TEST_READ_OBJ.wkg_hrs_list)

        self.assertEqual(n_months, 12)

    def test_num_months(self):
        """Check the number of months in the working hours file derived list."""

        n_months = len(TestBuilder.TEST_READ_OBJ.mth_list)

        self.assertEqual(n_months, 12)

    def test_num_months_span(self):
        """Check the number of months in the working hours file derived list."""

        n_months = len(TestBuilder.TEST_READ_OBJ.mth_span_list)

        self.assertEqual(n_months, 12)

    def test_only_integer(self):
        """Ensure only integer values exist."""

        for i in TestBuilder.TEST_READ_OBJ.wkg_hrs_list:

            self.assertIs(type(i), int)
