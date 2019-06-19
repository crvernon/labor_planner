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

    # expected keys in the YAML configuration file
    YAML_KEYS = ('input_directory', 'output_directory', 'staff_csv', 'work_hours_csv', 'fiscal_year')

    def __init__(self, config_file='/Users/d3y010/repos/github/labor_planning/config.yaml'):

        d = self.read_yaml(config_file)

        self.in_dir = self.check_directory(d['input_directory'])
        self.out_dir = self.check_directory(d['output_directory'])
        self.in_staff_csv = self.check_file(d['staff_csv'])
        self.in_work_hours = self.check_file(d['work_hours_csv'])

        # output files
        self.out_overview_file = "{0}/overview_chart.xlsx".format(self.out_dir)
        self.out_individ_file = "{0}/individual_staff_summary.xlsx".format(self.out_dir)
        self.out_rollup_file = "{0}/rollup.xlsx".format(self.out_dir)
        self.out_project_file = "{0}/projects.xlsx".format(self.out_dir)
        self.out_summary_file = "{0}/summary.xlsx".format(self.out_dir)

        self.fiscal_year = self.check_fiscal_year(d['fiscal_year'])

        # create data directory path where staff workbooks are stored
        self.data_dir = self.check_directory(os.path.join(self.in_dir, "FY_{}".format(self.fiscal_year)))

        # get last two digits of the fiscal year as a string
        self.fy = str(self.fiscal_year)[-2:]

        # get run design
        # TODO: add check for design type from acceptable options
        self.design = d['run_design']

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

        # ensure that all expected keys are present in the YAML dictionary
        missing_keys = set(ReadConfig.YAML_KEYS) - set(d.keys())

        if len(missing_keys) > 0:
            raise KeyError("Missing the keys {} in configuration file: {}".format(missing_keys, f))

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


if __name__ == '__main__':

    ini = '/Users/d3y010/repos/github/labor_planning/config.yaml'

    c = ReadConfig()

    print(c.fiscal_year)
