"""workbook_reader.py

Read and format staff labor planning workbooks.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""

import os
import collections

import numpy as np
import pandas as pd
import xlrd


class ReadWorkbooks:
    """Read in and process staff workbooks that contain hours per project in separate worksheets
     for all listed individuals.

    :param config_obj:                  YAML configuration object

    """

    def __init__(self, config_obj):

        self.my_settings = config_obj

        # set work hours and total month range rows
        self.work_hours_row, self.total_month_range_row, self.month_abbrev = self.get_work_hour_rows()

        # set month header row without total
        self.month_header = ['{}-{}'.format(i, config_obj.fy) for i in self.month_abbrev]

        # set month header row with total
        self.month_header_with_total = self.month_header + ['Total']

        # set source header value lists
        self.wkg_hours_hdr = self.work_hours_row[:-1]
        self.time_span_hdr = self.total_month_range_row

        # create list of input files
        self.file_list = self.get_files_list()

        # Create list of staff
        self.staff_list = self.get_staff_list()

        # Create list of months based on design
        self.month_list = self.create_time_span_list()

        # {staff_name: [hours per month, ...]}
        self.rollup_dict = {}

        # {project_number: project_title}
        self.prj_title_dict = {}

        # {staff_name: [hours per month for all projects, ...]}
        self.staff_dict = {}

        # {project: [project probability for all staff months, ...]}
        self.prj_prob_dict = {}

        # {staff_name: [hours associated with low probability funding, ...]}
        self.staff_low_prob_dict = {}

        # {staff_name: [hours associated with high probability funding, ...]}
        self.staff_high_prob_dict = {}

        # ordered dictionary of {staff_name: [[project_number, project_manager,
        #     total_hrs_per_mth, project_title, probability], ...}
        self.ind_dict = {}

        # {project_number:[[staff_name, project_manager, total_hrs_per_mth, probability], [...]], ...}
        self.projects_dict = {}

        # Iterate through directory xlsx files
        for f in self.file_list:

            # Get full path to workbook
            in_file = os.path.join(self.my_settings.data_dir, f)

            with xlrd.open_workbook(in_file) as wkbook:

                # create a list of worksheet names
                wksheets = wkbook.sheet_names()

                # See how many sheets there are
                wksheet_num = len(wksheets)

                # Create a list of index values to iterate through for sheets
                wksheet_idx_list = range(0, wksheet_num, 1)

                # Iterate through worksheets
                for index, i in enumerate(wksheet_idx_list):

                    # instantiate a worksheet
                    s = wkbook.sheet_by_name(wksheets[i])

                    # get all values in the column A
                    get_names = s.col_values(0)

                    # iterate through names in worksheet
                    for nm in get_names:

                        # if the name is not in the staff list, pass it
                        if nm in self.staff_list:

                            # add name to dict for monthly hour sum with placeholders
                            if nm not in self.rollup_dict:
                                self.rollup_dict[nm] = [0]*12

                            # get project number, proposal number, or work package number
                            prj_num = s.cell_value(rowx=2,colx=1)
                            prop_num = s.cell_value(rowx=2,colx=5)
                            wp_num = s.cell_value(rowx=2,colx=9)
                            mng_name = s.cell_value(rowx=8,colx=1)

                            # get project title - make no title if none listed
                            title = s.cell_value(rowx=3,colx=1)
                            if len(title) == 0:
                                title = 'No Title Listed'

                            # convert funding probability to decimal
                            fund_prob = self.set_probability(s.cell_value(rowx=7, colx=1))

                            # determine which project identifier to report
                            prj_id = self.get_prj_id(prj_num, prop_num, wp_num, nm, index)

                            # get project title from worksheet
                            if len(s.cell_value(3, 1)) == 0:
                                prj_title = 'none'
                            else:
                                prj_title = s.cell_value(3, 1)

                            # add project id and title to dictionary
                            if prj_id not in self.prj_title_dict:
                                self.prj_title_dict[prj_id] = prj_title

                            # get position of name in list by index
                            name_idx = get_names.index(nm)

                            # iterate through the hours estimate for each month
                            hrs_list = []
                            for m in self.month_list:
                                hrs = s.cell_value(rowx=name_idx,colx=m)

                                # only capture numeric values for hours, else 0
                                hr = self.check_int(hrs)

                                # create list of formatted hours for all probs
                                hrs_list.append(hr)

                                # add numeric hour values to a yearly hour list
                                self.staff_dict.setdefault(nm, []).append(hr)

                                # add project funding probability to dictionary
                                self.prj_prob_dict.setdefault(prj_id, []).append(fund_prob)

                                # differentiate between low and high funding probability
                                if fund_prob <= 0.5:
                                    self.staff_low_prob_dict.setdefault(nm, []).append(hr)
                                else:
                                    self.staff_high_prob_dict.setdefault(nm, []).append(hr)

                            # if hrs sum != 0 then add to project code to dict
                            if sum(hrs_list) != 0:
                                if nm not in self.ind_dict:
                                    self.ind_dict[nm] = [[prj_id, mng_name, hrs_list, title, fund_prob]]
                                else:
                                    self.ind_dict[nm].append([prj_id, mng_name, hrs_list, title, fund_prob])

                            # sum hours for each month per staff name for rollup workbook
                            #  project hours must be high prob of funding
                            if fund_prob > 0.5:
                                for ct, m_hr in enumerate(hrs_list):
                                    self.rollup_dict[nm][ct] = (self.rollup_dict[nm][ct] + m_hr)

        self.sort_staff_dict()

    @staticmethod
    def check_int(hour_value):
        """Checks to see if value is an integer and uses 0 if not"""
        try:
            hr = int(hour_value)
        except:
            hr = 0
        return hr

    @staticmethod
    def set_probability(p):
        """Set a probability between 0 and 1.

        :param p:                       Probability value from user.

        :return:                        Decimal probability from 0.0 to 1.

        """
        # convert funding probability to decimal
        if type(p) in (None, str):
            p = 100.0

        # account for fractional entries
        elif p <= 1:
            p *= 100.0

        return round(p / 100.0, 2)

    def get_work_hour_rows(self):
        """Read work hours file and generate the work hours and total month range rows.

        :param f:                       Full path with file name and extension to the work hours file.

        :return:                        [0] work hours row, [1] total month range row, [2] list of months abbrev.

        """
        df = pd.read_csv(self.my_settings.in_work_hours)

        # create work hours row
        hours_sum = df['work_hrs'].sum()
        work_hours_row = df['work_hrs'].tolist() + [hours_sum]

        # create total month range row
        df['month_row'] = df['start_mon'] + ' ' + df['start_day'].astype(str) + '-' + df['end_mon'] + ' ' + df['end_day'].astype(str)
        month_range_row = df['month_row'].tolist()

        # get a list of month names (abbreviations)
        month_abbrev_list = df['month']

        return work_hours_row, month_range_row, month_abbrev_list

    def get_files_list(self):
        """Generate a list of labor planning staff Excel files in the data directory.

        :return:                        List of labor planning files.

        """
        # acceptable Excel extensions
        extensions = ('.xlsx', '.xls')

        return [i for i in os.listdir(self.my_settings.data_dir) if os.path.splitext(i)[-1] in extensions]

    def get_staff_list(self):
        """Read and process input staff file.  A labor planning workbook will be generated for each
        staff member in this file.

        :return:                            List of staff full names

        """
        df = pd.read_csv(self.my_settings.in_staff_csv)

        df.fillna('', inplace=True)

        df['middle_initial'] = df['middle_initial'].replace(' ', '')

        df['full_name']= np.where(df['middle_initial'] == '',
                                     df['last_name'] + ', ' + df['first_name'],
                                     df['last_name'] + ', ' + df['first_name'])

        return df['full_name'].tolist()

    def create_time_span_list(self):
        """Create a list of 12 values to iterate through for col position.

        :return:                        List of months to include.

        """
        if self.my_settings.design == 'full_year':
            month_list = range(1, 13, 1)

        elif self.my_settings.design == 'quarter_2_3_4':
            month_list = range(1, 10, 1)

        elif self.my_settings.design == 'quarter_2_3':
            month_list = range(1, 7, 1)

        elif self.my_settings.design == 'quarter_2':
            month_list = range(1, 4, 1)

        elif self.my_settings.design == 'quarter_3_4_1':
            month_list = range(4, 13, 1)

        elif self.my_settings.design == 'quarter_3_4':
            month_list = range(4, 10, 1)

        return list(month_list)

    def type_tostring(self, in_val):
        """If value is number make int then string.

        """
        try:
            return str(int(in_val))

        except ValueError:
            return in_val

    def get_prj_id(self, project_number, proposal_number, task_number, staff_name, sheet_index):
        """Determine which project identifier to report"""

        if project_number != '':
            prj_id = self.type_tostring(project_number)

        elif project_number == '' and proposal_number != '':
            prj_id = self.type_tostring(proposal_number)

        elif project_number == '' and proposal_number == '' and task_number != '':
            prj_id = self.type_tostring(task_number)

        else:
            # format possible name issues
            s_name = staff_name.replace('*', '').strip()
            s_nospc = s_name.replace(' ', '_').replace('(', '').replace(')', '')
            sn = s_nospc.replace(',', '')
            # create reference name
            prj_id = "{0}_{1}".format(sn, sheet_index)

        return prj_id

    def sort_staff_dict(self):
        """Sort the staff dictionary by name"""

        # sort dict by staff name
        self.ind_dict = collections.OrderedDict(sorted(self.ind_dict.items()))

        # Iterate through the individual hours dictionary and reconfigure
        for k in self.ind_dict.keys():

            v = self.ind_dict[k]

            # format the worksheet name using modified staff name
            st_name = k.replace('*', '').strip()
            st_nospc = st_name.replace(' ', '_').replace('(', '').replace(')', '')
            s_name = st_nospc.replace(',', '')

            # iterate through each project for each staff member
            for i in v:
                prj = i[0]
                manager = i[1]
                hr_list = i[2]
                p_per = i[4]

                # add project id to dictionary
                if prj not in self.projects_dict:
                    self.projects_dict[prj] = [[s_name, manager, hr_list, p_per]]
                else:
                    self.projects_dict[prj].append([s_name, manager, hr_list, p_per])
