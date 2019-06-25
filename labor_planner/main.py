"""main.py

Execute labor planning code.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""
import argparse

from labor_planner.config_reader import ReadConfig
from labor_planner.workbook_reader import ReadWorkbooks
from labor_planner.stage_data import Stage
from labor_planner.labor_outputs.overview import Overview
from labor_planner.labor_outputs.project_level import Projects
from labor_planner.labor_outputs.individual_staff import IndividualHours
from labor_planner.labor_outputs.rollup_staff import Rollup
from labor_planner.labor_outputs.summary import Summary


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

    parser = argparse.ArgumentParser()
    parser.add_argument('config_file', type=str, help='Full path with file name to YAML configuration file.')
    args = parser.parse_args()

    LaborPlanner(args.config_file)
