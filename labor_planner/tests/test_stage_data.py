import os
import unittest

from labor_planner.config_reader import ReadConfig
from labor_planner.workbook_reader import ReadWorkbooks
from labor_planner.stage_data import Stage


class TestStageData(unittest.TestCase):
    """Test stage data attributes."""

    TEST_DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
    TEST_CONFIG_FILE = os.path.join(TEST_DATA_DIR, 'config_plan.yml')
    TEST_CONFIG_OBJ = ReadConfig(TEST_CONFIG_FILE)
    TEST_READ_OBJ = ReadWorkbooks(TEST_CONFIG_OBJ)
    TEST_DATA = Stage(TEST_CONFIG_OBJ, TEST_READ_OBJ)

    def test_avail_hours_sum(self):
        """Check available hours sum to ensure 0 or greater and less than 8784"""

        self.assertGreaterEqual(TestStageData.TEST_DATA.avail_hours_sum, 0)
        self.assertLessEqual(TestStageData.TEST_DATA.avail_hours_sum, 8784)

    def test_not_empty(self):
        """Ensure object not empty."""

        self.assertGreater(len(TestStageData.TEST_DATA.wkg_hours_hdr_list), 0)
        self.assertGreater(len(TestStageData.TEST_DATA.project_path_dict), 0)
        self.assertGreater(len(TestStageData.TEST_DATA.name_list), 0)
        self.assertGreater(len(TestStageData.TEST_DATA.percent_list), 0)

    def test_value_int(self):
        """Ensure attribute has an integer value."""

        self.assertIs(type(TestStageData.TEST_DATA.avail_hours_sum), int)
        self.assertIs(type(TestStageData.TEST_DATA.end_row), int)
