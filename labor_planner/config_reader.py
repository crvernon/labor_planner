"""config_reader.py

Configuration reader for YAML.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""

import os
import yaml


class ReadConfig:
    """Configuration reader for YAML.

    Read and validate an input YAML configuration file.

    :param config_file:         Full path with file name and extension to the input YAML configuration file.
    :type config_file:          str

    Attributes:
        in_dir (str):           Full path to the directory containing the staff Excel spreadsheets.
        out_dir (str):          Full path to the output directory.
        in_staff_csv (str):     Full path with file name and extension to the staff list CSV file.
        in_work_hours (str):    Full path with file name and extension to the work hours CSV file.
        fiscal_year (int):      Fiscal year in format YYYY

    """

    PROJECT_KEY = 'project'
    BUILDER_KEY = 'builder'
    PLANNER_KEY = 'planner'
    PROJECT_KEY_REQ = ['input_directory', 'staff_file', 'work_hours_csv', 'fiscal_year', 'staff_workbook_dir']
    BUILDER_KEY_REQ = ['num_blank_wksheets']
    PLANNER_KEY_REQ = ['output_directory', 'run_design']

    def __init__(self, config_file):

        d = self.read_yaml(config_file)

        # ensure all needed content is in the configuration file
        self.check_content(d)

        # project level settings
        project = d['project']

        self.in_dir = self.check_directory(project['input_directory'])
        self.in_staff_csv = self.check_file(project['staff_file'])
        self.in_work_hours = self.check_file(project['work_hours_csv'])
        self.fiscal_year = self.check_fiscal_year(project['fiscal_year'])
        self.data_dir = self.check_directory(project['staff_workbook_dir'])

        self.build = project['build_workbooks']
        self.plan = project['run_labor_planner']

        if (self.build is True) and (self.plan is True):
            msg = """`build_workbooks` and `run_labor_planner` cannot both be set to `True`.
                    Workbooks must be populated before running using `run_labor_planner` otherwise
                    the outputs will be empty."""

            raise RuntimeError(msg)

        if self.build:

            builder = d['builder']

            self.num_blank_wksheets = int(builder['num_blank_wksheets'])

        if self.plan:

            planner = d['planner']

            self.out_dir = self.check_directory(planner['output_directory'])
            self.design = planner['run_design']

            # output files
            self.out_overview_file = os.path.join(self.out_dir, "overview_chart.xlsx")
            self.out_individ_file = os.path.join(self.out_dir, "individual_staff_summary.xlsx")
            self.out_rollup_file = os.path.join(self.out_dir, "rollup.xlsx")
            self.out_project_file = os.path.join(self.out_dir, "projects.xlsx")
            self.out_summary_file = os.path.join(self.out_dir, "summary.xlsx")

        # get last two digits of the fiscal year as a string
        self.fy = str(self.fiscal_year)[-2:]

    @staticmethod
    def check_file(f):
        """Check the existence of a file.

        :param f:           Full path with file name and extension.
        :type f:            str

        :return:            Input file path and name with extension.
        """
        if os.path.isfile(f):
            return f
        else:
            raise FileNotFoundError(f)

    @staticmethod
    def check_directory(pth):
        """Check the existence of a file.

        :param pth:         Full path to the input directory.
        :type pth:          str

        :return:            Input directory path.
        """
        if os.path.dirname(pth):
            return pth
        else:
            raise NotADirectoryError(pth)

    @staticmethod
    def read_yaml(f):
        """Read and validate YAML configuration file as a dictionary.

        :param f:           Full path with file name and extension to YAML file.
        :type f:            str

        :return:            Dictionary.
        """
        with open(ReadConfig.check_file(f)) as yml:
            d = yaml.safe_load(yml)

        return d

    @staticmethod
    def check_fiscal_year(y):
        """Validate the format of the fiscal year.

        :param y:           Four digit year.
        :type y:            int

        :return             Four digit year.
        """
        if type(y) in (int, float):
            thous = y / 1000.0

            # ensure years are between 1900 and 2999
            if (thous >= 1) and (thous < 3):
                return int(y)
            else:
                raise ValueError("'fiscal_year' value {} not valid. Must be format YYYY with no quotes.".format(y))

        else:
            raise TypeError("'fiscal_year' value is type {}. Must be integer in YYYY format with no quotes.".format(y))

    @staticmethod
    def check_section_content(d, key, expected_list):
        """Ensure all desired parameters have been defined for a section

        :param d:                   Input config_obj object
        :param key:                 Category key
        :param expected_list:       List of subcategory keys

        """
        category_msg = 'Missing the following category key in configuration file:  {}'
        subcat_msg = 'Missing the following subcategories for the "{}" category in the configuration file:  {}'

        if key not in d:
            raise KeyError(category_msg.format(key))

        received_list = list(d[key].keys())

        key_valid = set(expected_list) - set(received_list)

        if len(key_valid) > 0:
            raise KeyError(subcat_msg.format(key, key_valid))

    @classmethod
    def check_content(cls, d):
        """Ensure all desired parameters have been defined in the configuration file.

        :param d:                   Input config_obj object

        """

        # validate project level settings
        cls.check_section_content(d, cls.PROJECT_KEY, cls.PROJECT_KEY_REQ)

        # validate build level settings
        if d[cls.PROJECT_KEY]['build_workbooks'] is True:
            cls.check_section_content(d, cls.BUILDER_KEY, cls.BUILDER_KEY_REQ)

        # validate planner level settings
        if d[cls.PROJECT_KEY]['run_labor_planner'] is True:
            cls.check_section_content(d, cls.PLANNER_KEY, cls.PLANNER_KEY_REQ)
