
import labor_planner.workbook_utils as util

from labor_planner.config_reader import ReadConfig
from labor_planner.workbook_reader import ReadWorkbooks
from labor_planner.stage_data import Stage
from labor_planner.overview_wkbook import Overview
from labor_planner.project_wkbook import Projects
from labor_planner.individual_hours_wkbook import IndividualHours
from labor_planner.rollup_wkbook import Rollup
from labor_planner.summary_wkbook import Summary


class LaborPlanner:

    def __init__(self, config):

        self.config_obj = ReadConfig(config)

        # get information from staff workbooks
        self.read_obj = ReadWorkbooks(self.config_obj)

        # stage data for labor analysis
        self.data = Stage(self.config_obj, self.read_obj)

        # build overview workbook
        Overview(self.config_obj, self.data)

        # build projects workbook
        Projects(self.config_obj, self.read_obj, self.data)

        # build individual hours workbook
        IndividualHours(self.config_obj, self.read_obj, self.data)

        # build rollup workbook
        Rollup(self.config_obj, self.read_obj, self.data)

        # build summary workbook
        Summary(self.config_obj, self.read_obj, self.data)


if __name__ == '__main__':

    config_file = '/Users/d3y010/repos/github/labor_planner/config.yaml'

    LaborPlanner(config_file)