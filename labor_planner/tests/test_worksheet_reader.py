import os
import unittest

from labor_planner.config_reader import ReadConfig
from labor_planner.workbook_reader import ReadWorkbooks


class TestWorksheetReader(unittest.TestCase):
    """Test configuration integrity."""

    TEST_DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
    TEST_CONFIG_FILE = os.path.join(TEST_DATA_DIR, 'config.yaml')
    TEST_CONFIG_OBJ = ReadConfig(TEST_CONFIG_FILE)
    TEST_READ_OBJ = ReadWorkbooks(TEST_CONFIG_OBJ)

    def test_num_months_total(self):
        """Check the number of months in the working hours file derived list."""

        n_months = len(TestWorksheetReader.TEST_READ_OBJ.month_abbrev)

        self.assertEqual(n_months, 12)

    def test_file_list_num(self):
        """Make sure there are files in the staff data directory to process."""

        n_files = len(TestWorksheetReader.TEST_READ_OBJ.file_list)

        self.assertGreater(n_files, 0)

    def test_rollup_dict(self):
        """Test that number of hours per month in staff sheets for each member is
        equal to the total number of months in the working hours file.

        """
        d = TestWorksheetReader.TEST_READ_OBJ.rollup_dict

        n_months = len(TestWorksheetReader.TEST_READ_OBJ.month_abbrev)

        # get number of staff in dictionary
        n_staff = len(d.keys())

        # get the number of months included for each staff member
        num_mth_per_staff = [len(d[k]) for k in d.keys()]

        # generate a comparable list
        ref_n_months = [n_months] * n_staff

        self.assertListEqual(num_mth_per_staff, ref_n_months)

    def test_prj_title_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.prj_title_dict), 0)

    def test_staff_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.staff_dict), 0)

    def test_prj_prob_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.prj_prob_dict), 0)

    def test_staff_low_prob_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.staff_low_prob_dict), 0)

    def test_staff_high_prob_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.staff_high_prob_dict), 0)

    def test_ind_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.ind_dict), 0)

    def test_project_dict(self):
        """Ensure dict not empty."""

        self.assertGreater(len(TestWorksheetReader.TEST_READ_OBJ.projects_dict), 0)


if __name__ == '__main__':

    unittest.main()
