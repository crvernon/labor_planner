import collections
import operator
import os
import sys
from datetime import date

import xlrd
import xlsxwriter
import easygui as eg
import pandas as pd

from labor_planning.config_reader import ReadConfig


class LaborSettings:

    def __init__(self, config):

        self.c = ReadConfig(config)

        # root dir where input, output, and data directories are located
        self.project_dir =  "/Users/d3y010/projects/organizational/labor_planning" #self.get_project_directory()

        # get design type
        self.design = "full_year" #self.get_design_type()

        # get the input data directory
        self.data_dir = "/Users/d3y010/projects/organizational/labor_planning/FY_2018" #self.get_data_directory()

        self.out_dir = "{0}/outputs".format(self.project_dir)
        self.in_staff_csv = "{0}/admin/staff_list.csv".format(self.project_dir)
        self.in_work_hrs_csv = "{0}/admin/work_hours.csv".format(self.project_dir)
        self.out_overview_file = "{0}/overview_chart.xlsx".format(self.out_dir)
        self.out_individ_file = "{0}/individual_staff_summary.xlsx".format(self.out_dir)
        self.out_rollup_file = "{0}/rollup.xlsx".format(self.out_dir)
        self.out_project_file = "{0}/projects.xlsx".format(self.out_dir)
        self.out_summary_file = "{0}/summary.xlsx".format(self.out_dir)
        self.fiscal_year = int(self.data_dir.split('.')[0][-4:])
        self.fy = self.data_dir.split('.')[0][-4:]

    def get_project_directory(self):
        """Get the project directory where the data directories are kept (e.g., FY2018, etc.).
        This allows the user to either choose the default directory or select a different one.

        :return:                            The target project directory.

        """
        default_dir_msg = "Do you wish to use the default project directory: {0}".format(self.c.in_dir)
        select_dir_msg = "Select the project directory of your data:"
        title = "Labor Planning Project Directory"

        # get user input
        yn = eg.ynbox(default_dir_msg, title)

        if yn is False:
            return eg.diropenbox(select_dir_msg, title)

        else:
            return self.c.in_dir

    @staticmethod
    def get_design_type():
        """
        Select the design of the output reports.  The following designs are available:

        `Full Year`; `Quarters 2, 3`; `Quarters 2, 3, 4`; `Quarters 3, 4, 1`;
        `Quarters 3, 4`; and `Quarters 4, 1`

        :return:                        The design type variable from user choice.

        """
        # design type as selected by the user to the look up value
        d = {'Full Year': 'full_year',
                'Quarters 2, 3': 'quarter_2_3',
                'Quarters 2, 3, 4': 'quarter_2_3_4',
                'Quarters 3, 4, 1': 'quarter_3_4_1',
                'Quarters 3, 4': 'quarter_3_4',
                'Quarters 4, 1': 'quarter_4_1'}
        
        # text for gui
        msg = "Select the target design:"
        title = 'Labor Planning Setup'
        options = list(d.keys())

        # get user input
        return d[eg.choicebox(msg, title, options)]

    def get_data_directory(self):
        """
        Select the directory in which the labor planning excel files are stored. 
        The directory needs to be named using the following format:  `FY2018` where 2018 is 
        the fiscal year that is to be processed.
        
        
        :return:                        The data directory chosen by the user.
        """
        msg = "Select the data directory you wish to process: "
        title = 'Labor Planning Directory'
        options = [os.path.join(self.project_dir, i) for i in os.listdir(self.project_dir) if 'FY' in i]

        return eg.choicebox(msg, title, options)


def type_tostring(in_val):
    """If value is number make int then string.

    """
    try:
        return str(int(in_val))

    except ValueError:
        return in_val


def get_prj_id(project_number, proposal_number, task_number, staff_name, sheet_index):
    """Determine which project identifier to report"""
    if project_number != '':
        prj_id = type_tostring(project_number)
    elif project_number == '' and proposal_number != '':
        prj_id = type_tostring(proposal_number)
    elif project_number == '' and proposal_number == '' and task_number != '':
        prj_id = type_tostring(task_number)
    else:
        # format possible name issues
        s_name = staff_name.replace('*','').strip()
        s_nospc = s_name.replace(' ','_').replace('(','').replace(')','')
        sn = s_nospc.replace(',','')
        # create reference name
        prj_id = "{0}_{1}".format(sn, sheet_index)
    return prj_id


def check_int(hour_value):
    """Checks to see if value is an integer and uses 0 if not"""
    try:
        hr = int(hour_value)
    except:
        hr = 0
    return hr


def implement_design(design, in_list):
    """Gets a portion of a list that is based on the specified design"""
    if design == 'full_year':
        out_list = in_list
    elif design == 'quarter_2_3_4':
        out_list = in_list[0:9]
    elif design == 'quarter_2_3':
        out_list = in_list[0:6]
    elif design == 'quarter_2':
        out_list = in_list[0:3]
    elif design == 'quarter_3_4_1':
        out_list = in_list[-9:]
    elif design == 'quarter_3_4':
        out_list = in_list[3:9]
    return out_list


def get_nested_list_sum(in_list, list_index):
    """Returns the sum of a nested list"""
    out_list = sum([sum(i[list_index]) for i in in_list])
    return out_list


def open_xlsx(in_file):
    # open workbook
    try:
        # Open the workbook
        return xlrd.open_workbook(in_file)
    except:
        #print "ERROR:  The following file cannot be opened: {0}".format(in_file)
        #print "Check to see if the file is opened by another user."
        #print "Run again after workbook has been closed."
        #print "Exiting..."
        sys.exit()


def enable_formatting(in_workbook_instance):
    """
    USAGE:  Index options for preconfigured formatting:
        [0]  align_center
        [1]  align_right
        [2]  bold_1
        [3]  bold_1_center
        [4]  bold_big
        [5]  bold_right
        [6]  border_1
        [7]  border_gray
        [8]  border_gray_right
        [9]  border_gray_center
        [10] num_percent
        [11] url
        [12] merge_header
        [13] merge_border
        [14] bkg_red
        [15] bkg_yellow
        [16] bkg_green
        [17] bkg_red_bold
        [18] bkg_yellow_bold
        [19] bkg_green_bold
    """
    # Formatting options
    align_center = in_workbook_instance.add_format({'align': 'center'})
    align_right = in_workbook_instance.add_format({'align': 'right'})
    bold_1 = in_workbook_instance.add_format({'bold': 1})
    bold_1_center = in_workbook_instance.add_format({'bold': 1, 'align': 'center'})
    bold_big = in_workbook_instance.add_format({'bold': 2, 'size': 22})
    bold_right = in_workbook_instance.add_format({'bold': 1,'align': 'right'})
    border_1 = in_workbook_instance.add_format({'border': 1})
    border_gray = in_workbook_instance.add_format({'border': 1, 'bg_color': '#BDBDBD'})
    border_gray_right = in_workbook_instance.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'right'})
    border_gray_center = in_workbook_instance.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'center'})
    num_percent = in_workbook_instance.add_format({'num_format': '0%'})
    url = in_workbook_instance.add_format({'font_color': 'blue', 'underline': 0})
    merge_header = in_workbook_instance.add_format({'bold': 1, 'align': 'center', 'bottom': 1})
    merge_border = in_workbook_instance.add_format({'border': 1, 'align': 'center', 'bold': 1})
    bkg_red = in_workbook_instance.add_format({'bg_color': '#F5A9BC'})
    bkg_yellow = in_workbook_instance.add_format({'bg_color': '#F2F5A9'})
    bkg_green = in_workbook_instance.add_format({'bg_color': '#04B404'})
    bkg_red_bold = in_workbook_instance.add_format({'bold': 1, 'bg_color': '#F5A9BC', 'align': 'center'})
    bkg_yellow_bold = in_workbook_instance.add_format({'bold': 1, 'bg_color': '#F2F5A9', 'align': 'center'})
    bkg_green_bold = in_workbook_instance.add_format({'bold': 1, 'bg_color': '#04B404', 'align': 'center'})
    # Return options
    return align_center, align_right, bold_1, bold_1_center, bold_big, \
            bold_right, border_1, border_gray, border_gray_right, \
            border_gray_center, num_percent, url, merge_header, merge_border, \
            bkg_red, bkg_yellow, bkg_green, bkg_red_bold, bkg_yellow_bold, \
            bkg_green_bold


def set_merge_range(worksheet_instance, design, fiscal_year, start_row_num, format_type):
    # Make two digit year
    yr = int(str(fiscal_year)[2:4])
    # Set design specific criteria
    if design == 'full_year':
        # set totals and percent column names
        totals_column = 'N'
        percent_column = 'O'
        # define date range for design
        date_range = 'Dec 27, {0} - Dec 25, {1}'.format(fiscal_year, (fiscal_year + 1))
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 2 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('E{0}:G{0}'.format(start_row_num), 'Quarter 3 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('H{0}:J{0}'.format(start_row_num), 'Quarter 4 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('K{0}:M{0}'.format(start_row_num), 'Quarter 1 - FY{0}'.format(yr + 1), format_type)
    elif design == 'quarter_2_3_4':
        # set totals column name
        totals_column = 'K'
        percent_column = 'L'
        # define date range for design
        date_range = 'Dec 27, {0} - Sep 30, {0}'.format(fiscal_year)
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 2 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('E{0}:G{0}'.format(start_row_num), 'Quarter 3 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('H{0}:J{0}'.format(start_row_num), 'Quarter 4 - FY{0}'.format(yr), format_type)
    elif design == 'quarter_2_3':
        # set totals column name
        totals_column = 'H'
        percent_column = 'I'
        # define date range for design
        date_range = 'Dec 27, {0} - Jun 26, {0}'.format(fiscal_year)
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 2 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('E{0}:G{0}'.format(start_row_num), 'Quarter 3 - FY{0}'.format(yr), format_type)
    elif design == 'quarter_2':
        # set totals column name
        totals_column = 'E'
        percent_column = 'F'
        # define date range for design
        date_range = 'Dec 27, {0} - Mar 27, {0}'.format(fiscal_year)
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 2 - FY{0}'.format(yr), format_type)
    elif design == 'quarter_3_4_1':
        # set totals column name
        totals_column = 'K'
        percent_column = 'L'
        # define date range for design
        date_range = 'May 28, {0} - Dec 25, {0}'.format(fiscal_year)
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 3 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('E{0}:G{0}'.format(start_row_num), 'Quarter 4 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('H{0}:J{0}'.format(start_row_num), 'Quarter 1 - FY{0}'.format(yr+1), format_type)
    elif design  == 'quarter_3_4':
        # set totals column name
        totals_column = 'H'
        percent_column = 'I'
        # define date range for design
        date_range = 'May 28, {0} - Sep 30, {0}'.format(fiscal_year)
        # write merged headers
        worksheet_instance.merge_range('B{0}:D{0}'.format(start_row_num), 'Quarter 3 - FY{0}'.format(yr), format_type)
        worksheet_instance.merge_range('E{0}:G{0}'.format(start_row_num), 'Quarter 4 - FY{0}'.format(yr), format_type)
    return totals_column, percent_column, date_range

# Sort ind_dict by staff name
def sort_dict(d):
    return collections.OrderedDict(sorted(d.items()))

def get_total_project_hours(sheet, totals_col_index):
    vals = sheet.col_values(totals_col_index)
    #print vals


def get_design_type(txt):
    d = {'Full Year': 'full_year',
            'Quarters 2, 3': 'quarter_2_3',
            'Quarters 2, 3, 4': 'quarter_2_3_4',
            'Quarters 3, 4, 1': 'quarter_3_4_1',
            'Quarters 3, 4': 'quarter_3_4',
            'Quarters 4, 1': 'quarter_4_1'}
    return d[txt]

def get_year():
    return date.today().year


def check_exist(data_dir):
    """Check the existence of a directory.

    :param directory:                       Full path to a directory.

    :return:                                Boolean response
    """
    if not os.path.isdir(data_dir):
        return False
    else:
        return True



def check_file_exist(f):
    """Check the existence of a file.  Pop-up box if not found.

    :param f:               Full path with file name and extension to input file.

    """
    # check existence of the data directory
    if os.path.isfile(f) is False:
        msg = "File '{}' does not exist.".format(f)
        title = 'Error'
        button_text = 'EXIT'

        eg.msgbox(msg, title, button_text)

        raise FileNotFoundError



###############################################################################
# -- RUN --
###############################################################################

ini = '/Users/d3y010/repos/github/labor_planning/config.yaml'

# get settings
my_settings = LaborSettings(ini)

# check existence of staff list
check_file_exist(my_settings.in_staff_csv)

# blank objects
staff_dict = {}
staff_low_prob_dict = {}
staff_high_prob_dict = {}
ind_dict = {}
prj_prob_dict = {}
rollup_dict = {}
project_path_dict = {}
projects_dict = {}
prj_title_dict = {}
check_dups_list = []


class ReadWorkbooks:

    def __init__(self, my_settings):

        self.my_settings = my_settings

        self.month_abbrev = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        # set month header row without total
        self.month_header = ['{}-{}'.format(i, my_settings.fy) for i in self.month_abbrev]

        # set month header row with total
        self.month_header_with_total = self.month_header + ['Total']

        # set work hours and total month range rows
        self.work_hours_row, self.total_month_range_row = self.get_work_hour_rows()

        # set source header value lists
        self.wkg_hours_hdr = self.work_hours_row[:-1]
        self.time_span_hdr = self.total_month_range_row[:-1]

###############################################################################

        # Update user
        #print "\nProcessing files..."
        #eg.message_box("Processing Data...", "UPDATE", "NEXT")

        # create list of input files
        self.file_list = self.get_files_list()

        # Create list of staff
        self.staff_list = self.get_staff_list()

        # Create list of months based on design
        self.month_list = self.create_time_span_list()


        # TODO:  START HERE

        # Iterate through directory xlsx files
        for f in self.file_list:

            # Get full path to workbook
            in_file = "{0}/{1}".format(self.my_settings.data_dir, f)

            # open workbook
            wkbook = open_xlsx(in_file)

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

                # make all sheets active
                sheet_sum_cell = 1

                # if total hours is 0 then pass the sheet
                if sheet_sum_cell == 0:
                    pass

                else:
                    # get all values in the column A
                    get_names = s.col_values(0)

                    # iterate through names in worksheet
                    for nm in get_names:

                        # if the name is not in the staff list, pass it
                        if nm in self.staff_list:

                            # add name to dict for monthly hour sum with placeholders
                            if nm not in rollup_dict:
                                rollup_dict[nm] = [0]*12

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
                            p = s.cell_value(rowx=7,colx=1)
                            if type(p) in (None, str, unicode):
                                p = 100

                            if p == 1:
                                p = 100
                            elif p < 1:
                                p = p * 100

                            try:
                                fund_prob = round(p/100, 2)
                            except TypeError:
                                raise(TypeError("Probability '{}' is string instead of number for file {} on sheet {}".format(type(p), f, wksheets[i])))

                            # determine which project identifier to report
                            prj_id = get_prj_id(prj_num, prop_num, wp_num, nm, index)

                            # get project title from worksheet
                            if len(s.cell_value(3, 1)) == 0:
                                prj_title = 'none'
                            else:
                                prj_title = s.cell_value(3, 1)

                            # add project id and title to dictionary
                            if prj_id not in prj_title_dict:
                                prj_title_dict[prj_id] = prj_title
                            else:
                                pass

                            # get position of name in list by index
                            name_idx = get_names.index(nm)

                            # get the value in the cell name
                            cell_val = s.cell_value(rowx=name_idx,colx=0)

                            # iterate through the hours estimate for each month
                            hrs_list = []
                            for m in self.month_list:
                                hrs = s.cell_value(rowx=name_idx,colx=m)

                                # only capture numeric values for hours, else 0
                                hr = check_int(hrs)

                                # create list of formatted hours for all probs
                                hrs_list.append(hr)

                                # add numeric hour values to a yearly hour list
                                staff_dict.setdefault(nm, []).append(hr)

                                # add project funding probability to dictionary
                                prj_prob_dict.setdefault(prj_id, []).append(fund_prob)

                                # differentiate between low and high funding probability
                                if fund_prob <= 0.5:
                                    staff_low_prob_dict.setdefault(nm, []).append(hr)
                                else:
                                    staff_high_prob_dict.setdefault(nm, []).append(hr)

                            # if hrs sum != 0 then add to project code to dict
                            if sum(hrs_list) != 0:
                                if nm not in ind_dict:
                                    ind_dict[nm] = [[prj_id, mng_name, hrs_list, title, fund_prob]]
                                else:
                                    ind_dict[nm].append([prj_id, mng_name, hrs_list, title, fund_prob])

                            # sum hours for each month per staff name for rollup workbook
                            #  project hours must be high prob of funding
                            if fund_prob > 0.5:
                                for ct, m_hr in enumerate(hrs_list):
                                    rollup_dict[nm][ct] = (rollup_dict[nm][ct] + m_hr)


    def get_work_hour_rows(self):
        """Read work hours file and generate the work hours and total month range rows.

        :param f:                       Full path with file name and extension to the work hours file.

        :return:                        [0] work hours row, [1] total month range row

        """
        df = pd.read_csv(self.my_settings.in_work_hrs_csv)

        # create work hours row
        hours_sum = df['work_hrs'].sum()
        work_hours_row = df['work_hrs'].tolist() + [hours_sum]

        # create total month range row
        df['month_row'] = df['start_mon'] + ' ' + df['start_day'].astype(str) + '-' + df['end_mon'] + ' ' + df['end_day'].astype(str)
        month_range_row = df['month_row'].tolist()

        return work_hours_row, month_range_row

    def get_files_list(self):
        """Generate a list of labor planning staff Excel files in the data directory.

        :return:                        List of labor planning files.

        """
        # acceptable Excel extensions
        extensions = ('.xlsx', '.xls')

        return [i for i in os.listdir(self.my_settings.data_dir) if os.path.splitext(i)[-1] in extensions]

    def get_staff_list(self):
        """Get a list of staff from the current staff file.

        :return:                        List of staff to evaluate.

        """
        with open(my_settings.in_staff_csv) as get:
            s = get.read()

        return [i for i in s.split(';') if len(i) > 0]


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



ReadWorkbooks(my_settings)



################################################################################
# sort dict by staff name
ind_dict = sort_dict(ind_dict)
# Iterate through the individual hours dictionary and reconfigure
for k, v in ind_dict.iteritems():
    # format the worksheet name using modified staff name
    st_name = k.replace('*','').strip()
    st_nospc = st_name.replace(' ','_').replace('(','').replace(')','')
    s_name = st_nospc.replace(',','')
    # iterate through each project for each staff member
    for i in v:
        # assign values
        prj = i[0]
        manager = i[1]
        hr_list = i[2]
        p_per = i[4]
        # add project id to dictonary
        if prj not in projects_dict:
            projects_dict[prj] = [[s_name, manager, hr_list, p_per]]
        else:
            projects_dict[prj].append([s_name, manager, hr_list, p_per])

################################################################################

# Set header content based on design
wkg_hours_hdr_list = implement_design(my_settings.design, self.wkg_hours_hdr)

# Calculate available hours sum
avail_hours_sum = sum(wkg_hours_hdr_list)

################################################################################
# Create new list removing **Post MA assessments, keep them in a separate list
post_ma_list = []
combine_list = []
for k, v in staff_dict.iteritems():
    if '**' in k:
        post_ma_list.append(sum(v))
    else:
        #-- calculate percent covered based on total hours/ avail work hours
        total_percent_covered = round((float(sum(v))/float(avail_hours_sum)),2)
        combine_list.append([k, total_percent_covered])

# Sort by last name ascending
combine_list.sort()

# Create separte list for names and percent covered
name_list = [ i[0] for i in combine_list ]
percent_list = [ i[1] for i in combine_list ]

# Return post masters list
post_masters_hours = sum(post_ma_list)

################################################################################
# Create high probability lists
high_prob_list = []
for k, v in staff_high_prob_dict.iteritems():
    if '**' in k:
        pass
    else:
        high_prob_percent = round((float(sum(v))/float(avail_hours_sum)),2)
        high_prob_list.append([k, high_prob_percent])

# Sort by last name ascending
high_prob_list.sort()

# Create separte list for names and percent covered
high_name_list = [ i[0] for i in high_prob_list ]
high_percent_list = [ i[1] for i in high_prob_list ]

################################################################################
# Create low probability lists
low_prob_list = []
for k, v in staff_low_prob_dict.iteritems():
    if '**' in k:
        pass
    else:
        low_prob_percent = round((float(sum(v))/float(avail_hours_sum)),2)
        low_prob_list.append([k, low_prob_percent])

# Sort by last name ascending
low_prob_list.sort()

# Create separte list for names and percent covered
low_name_list = [ i[0] for i in low_prob_list ]
low_percent_list = [ i[1] for i in low_prob_list ]

################################################################################
# Create lists to populate output chart values
look_up_list = []
for i in name_list:
    if i in high_name_list and i in low_name_list:
        look_up_list.append([i, high_percent_list[high_name_list.index(i)], \
                                low_percent_list[low_name_list.index(i)]] )
    elif i in high_name_list and i not in low_name_list:
        look_up_list.append([i, high_percent_list[high_name_list.index(i)], 0.0])
    elif i in low_name_list and i not in high_name_list:
        look_up_list.append([i, 0.0, low_percent_list[low_name_list.index(i)]])
    else:
        look_up_list.append([i, 0.0, 0.0])

# Sort look up list
look_up_list.sort()

# Create lists for outputs
full_name_list = [ i[0] for i in look_up_list ]
full_high_list = [ i[1] for i in look_up_list ]
full_low_list = [ i[2] for i in look_up_list ]

# Get value for the length of name column
end_row = (len(full_name_list) + 1)


# update user
eg.message_box("Creating output workbooks...", "UPDATE", "NEXT")






################################################################################
# Create overview chart workbook
################################################################################
# Create graph Excel workbook and instantiate worksheet
graph_workbook = xlsxwriter.Workbook(my_settings.out_overview_file)
graph_ws1 = graph_workbook.add_worksheet()

# Create formatting objects
header_merge_format = enable_formatting(graph_workbook)[12]
bold_1 = enable_formatting(graph_workbook)[2]
bold_center = enable_formatting(graph_workbook)[3]
url_format = enable_formatting(graph_workbook)[11]

# List of headers
merged_header_text = 'Proportion of FTE Covered in FY{0}'.format(my_settings.fy)
header_list = ['> 50% Funding Probability', '<= 50% Funded Probability', 'All Projects']

# Set hover over information
hyperlink_tip = 'Click name to open source workbook.'

# Set column widths
graph_ws1.set_column('A:A', 20)
graph_ws1.set_column('B:D', 28)

# Write headers to worksheet
graph_ws1.merge_range('B1:D1', 'Proportion of FTE Covered in FY{0}'.format(my_settings.fy), header_merge_format)
graph_ws1.write('A2', 'Staff Member', bold_1)
graph_ws1.write_row('B2', header_list, bold_center)

# Write data to worksheet
graph_ws1.write_column('A3', full_name_list)
graph_ws1.write_column('B3', full_high_list)
graph_ws1.write_column('C3', full_low_list)
graph_ws1.write_column('D3', percent_list)

# Create hyperlinks
# --create hyperlink for each staff member with an existing workbook
# --start at A3, idx = 3
idx = 3
for i in full_name_list:
    # set hyperlink parameters
    cell_location = "A{0}".format(idx)

    # format possible name issues
    l_name = i.replace('*','').strip()
    l_nospc = l_name.replace(' ','_').replace('(','').replace(')','')
    l = l_nospc.replace(',','')

    # create link location
    link_location = "{0}#{1}!a1".format(os.path.basename(my_settings.out_individ_file), l)

    # create hyperlink if link location exists
    graph_ws1.write_url(cell_location, link_location, url_format, i, hyperlink_tip)

    # advance index for cell location
    idx += 1

# Create chart object
chart = graph_workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

# Add total percent covered series
chart.add_series({
    'name':         'Prob > 50%',
    'categories':   '=Sheet1!$A$3:$A${0}'.format(end_row+1),
    'values':       '=Sheet1!$B$3:$B${0}'.format(end_row+1)
    })

# Add probable (<= 50%) percent covered series
chart.add_series({
    'name':         'Prob <= 50%',
    'categories':   '=Sheet1!$A$3:$A${0}'.format(end_row+1),
    'values':       '=Sheet1!$C$3:$C${0}'.format(end_row+1)
    })

# Set chart style and size
chart.set_style(18)
chart.set_size({'x_scale': 2.5, 'y_scale': 2})
chart.set_title({'name': 'Estimated Project Hours'})
chart.set_y_axis({'name': 'Proportion of FTE'})

# Add chart to worksheet
graph_ws1.insert_chart('F3', chart)

# Close workbook
graph_workbook.close()

################################################################################
# Create projects workbook
################################################################################
# Create projects Excel workbook and instantiate worksheet
project_wkbook = xlsxwriter.Workbook(my_settings.out_project_file)

# Create formatting objects
bold_1 = enable_formatting(project_wkbook)[2]
big_bold = enable_formatting(project_wkbook)[4]
border_1 = enable_formatting(project_wkbook)[6]
merge_format = enable_formatting(project_wkbook)[13]
border_gray = enable_formatting(project_wkbook)[7]
border_gray_right = enable_formatting(project_wkbook)[8]
border_gray_center = enable_formatting(project_wkbook)[9]
percent_format = enable_formatting(project_wkbook)[10]

# Set worksheet index for naming convention
ws_index = 0

# Iterate through project dictionary
for k, v in projects_dict.iteritems():

    # get project manager name
    pm = v[0][1]
    pb = v[0][3]

    # instantiate worksheet
    set_ws_name = 'sheet_{0}'.format(ws_index)
    prj_ws = project_wkbook.add_worksheet(set_ws_name)
    ws_index += 1

    # add project id and path to sheet to dictionary
    if k not in project_path_dict:
        prj_id_path = "{0}#{1}!a1".format(os.path.basename(my_settings.out_project_file), set_ws_name)
        project_path_dict[k] = prj_id_path
    else:
        print("ERROR:  Duplicate project ids detected for: ", k)

    # Set column widths
    prj_ws.set_column('A:A', 20)
    prj_ws.set_column('B:M', 24)

    # Set header content based on design
    month_hdr_list = implement_design(my_settings.design, month_hdr)
    wkg_hours_hdr_list = implement_design(my_settings.design, self.wkg_hours_hdr)
    time_span_hdr_list = implement_design(my_settings.design, self.time_span_hdr)

    # Calculate available hours sum
    avail_hours_sum = sum(wkg_hours_hdr_list)

    # Set design specific criteria
    des_inst = set_merge_range(prj_ws, my_settings.design, my_settings.fiscal_year, 9, merge_format)
    totals_column = des_inst[0]
    percent_column = des_inst[1]
    date_range = des_inst[2]

    # Write generic sheet content
    prj_ws.write('A1', 'Staff Planning', big_bold)
    prj_ws.write('A3', 'Project Rollup', bold_1)
    prj_ws.write_row('A4', ['Project ID:', k])
    prj_ws.write_row('A5', ['Project Title:', prj_title_dict[k]])
    prj_ws.write_row('A6', ['Project Manager:', pm])
    prj_ws.write_row('A7', ['Funding Probability:', pb])
    prj_ws.write('A9', 'Technical Group', border_1)
    prj_ws.write('A10', '', border_gray)
    prj_ws.write('A11', 'Wkg Hrs Available =', border_gray_right)
    prj_ws.write('A12', 'Processing Month =', border_gray_right)

    # Write headers based on design
    prj_ws.write_row('B10', month_hdr_list, border_gray_center)
    prj_ws.write_row('B11', wkg_hours_hdr_list, border_gray_center)
    prj_ws.write_row('B12', time_span_hdr_list, border_gray_center)
    prj_ws.write('{0}10'.format(totals_column), 'Total', border_gray_center)
    prj_ws.write('{0}11'.format(totals_column), avail_hours_sum, border_gray_center)
    prj_ws.write('{0}12'.format(totals_column), date_range, border_gray_center)
    prj_ws.write('{0}10'.format(percent_column), '', border_gray_center)
    prj_ws.write('{0}11'.format(percent_column), 'Percent Covered', border_gray_center)
    prj_ws.write('{0}12'.format(percent_column), '', border_gray_center)

    # Set start cell position
    cell_position = 13

    # Iterate through each person working on a project
    for i in v:
        # assign variables
        staff_name = i[0]
        staff_hours = i[2] #implement_design(design, i[2])
        staff_hours_sum = sum(staff_hours)
        percent_covered = (float(staff_hours_sum) / avail_hours_sum)
        # write variables to sheet
        prj_ws.write('A{0}'.format(cell_position), staff_name)
        prj_ws.write_row('B{0}'.format(cell_position), staff_hours)
        # write total hours
        prj_ws.write('{0}{1}'.format(totals_column, cell_position), staff_hours_sum)
        # write percent covered
        prj_ws.write('{0}{1}'.format(percent_column, cell_position), percent_covered, percent_format)
        # advance cell position
        cell_position += 1

# Close rollup workbook
project_wkbook.close()

################################################################################
# Create individual hours planning Excel workbook
################################################################################
indiv_plan_wkbook = xlsxwriter.Workbook(my_settings.out_individ_file)

# Create formatting objects
right_bold = indiv_plan_wkbook.add_format({'bold': 1,'align': 'right'})
bold_1 = indiv_plan_wkbook.add_format({'bold': 1})
big_bold = indiv_plan_wkbook.add_format({'bold':2, 'size': 22})
merge_format = indiv_plan_wkbook.add_format({'border': 1, 'align': 'center'})
center = indiv_plan_wkbook.add_format({'align': 'center'})
right = indiv_plan_wkbook.add_format({'align': 'right'})
url_format = indiv_plan_wkbook.add_format({'font_color': 'blue', 'underline': 0})

# Set hover over information
hyperlink_tip = 'Click name to open project workbook.'

# Iterate through the individual hours dictionary and write content to file
for k, v in ind_dict.iteritems():
    # format the worksheet name using modified staff name
    st_name = k.replace('*','').strip()
    st_nospc = st_name.replace(' ','_').replace('(','').replace(')','')
    ws_name = st_nospc.replace(',','')

    # instantiate worksheet
    ws = indiv_plan_wkbook.add_worksheet(ws_name)

    # set column E and total column by design
    if my_settings.design == 'quarter_3_4':
        row_13_l = self.month_header_with_total[3:9]
        row_13_l.append(self.month_header_with_total[-1])
        row_14_l = self.work_hours_row[3:9]
        row_14_l.append(sum(row_14_l))
        row_15_l = self.total_month_range_row[3:9]
        row_15_l.append(self.total_month_range_row[-1])
        end_col = 'J'
        total_col = 'K'
    elif my_settings.design == 'quarter_3_4_1':
        row_13_l = self.month_header_with_total[-10:]
        row_14_l = self.work_hours_row[3:12]
        row_14_l.append(sum(row_14_l))
        row_15_l = self.total_month_range_row[-10:]
        end_col = 'M'
        total_col = 'N'
    elif my_settings.design == 'full_year':
        row_13_l = self.month_header_with_total[-1:]
        row_14_l = self.work_hours_row[0:12]
        row_14_l.append(sum(row_14_l))
        row_15_l = self.total_month_range_row[0:12]
        end_col = 'P'
        total_col = 'Q'
    elif my_settings.design == 'quarter_2_3':
        row_13_l = self.month_header_with_total[:6]
        row_13_l.append(self.month_header_with_total[-1])
        row_14_l = self.work_hours_row[:6]
        row_14_l.append(sum(row_14_l))
        row_15_l = self.total_month_range_row[:6]
        row_15_l.append(self.total_month_range_row[-1])
        end_col = 'J'
        total_col = 'K'
    elif my_settings.design == 'quarter_2_3_4':
        row_13_l = self.month_header_with_total[:9]
        row_14_l = self.work_hours_row[:9]
        row_14_l.append(sum(row_14_l))
        row_15_l = self.total_month_range_row[:9]
        end_col = 'M'
        total_col = 'N'

    # set column widths
    ws.set_column('A:A', 10)
    ws.set_column('B:B', 10)
    ws.set_column('C:C', 40)
    ws.set_column('D:D', 20)
    ws.set_column('E:{0}'.format(end_col), 15)
    ws.set_column('{0}:{0}'.format(total_col), 30)

    # write column A information
    ws.write('A1', 'Staff Planning', big_bold)
    ws.write('A3', 'Staff Name:', right_bold)
    ws.write('A5', 'PROJECT', bold_1)
    ws.write('A6', 'NUMBER', bold_1)

    # write column B information
    ws.write('B5', 'FUNDING', bold_1)
    ws.write('B6', 'PROBABILITY', bold_1)

    # write column C information
    ws.write('C5', 'PROJECT', bold_1)
    ws.write('C6', 'DESCRIPTION', bold_1)

    # write column D information
    ws.write('B3', '{0}'.format(st_name), big_bold)
    ws.write('D5', 'MANAGER or', bold_1)
    ws.write('D6', 'TASK MANAGER', bold_1)
    ws.write('D7', 'Wkg Hrs Available =', right)
    ws.write('D8', 'Processing Month =', right)

    # write column E information
    ws.write_row('E6', row_13_l, center)
    ws.write_row('E7', row_14_l, center)
    ws.write_row('E8', row_15_l, center)

    # create merged range cell values
    if my_settings.design == 'quarter_3_4':
        ws.merge_range('E5:G5', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('H5:J5', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
    elif my_settings.design == 'quarter_2_3':
        ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
    elif my_settings.design == 'quarter_3_4_1':
        ws.merge_range('E5:G5', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('H5:J5', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('K5:M5', 'Quarter 1 - FY{0}'.format(int(my_settings.fy)+1), merge_format)
    elif my_settings.design == 'quarter_2_3_4':
        ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('K5:M5', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
    elif my_settings.design == 'full_year':
        ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('K5:M5', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
        ws.merge_range('N5:P5', 'Quarter 1 - FY{0}'.format(int(my_settings.fy)+1), merge_format)

    #ws.merge_range('L5:N5', 'Quarter 1 - FY17', merge_format)

    # write content to worksheet
    for iteration, i in enumerate(v):
        try:
            # create hyperlink if link location exists for project id
            link_location = project_path_dict[i[0]]
            ws.write_url('A{0}'.format(iteration+9), link_location, url_format, i[0], hyperlink_tip)
            # write project funding probability
            ws.write('B{0}'.format(iteration+9), i[4])
            # write project name
            ws.write('C{0}'.format(iteration+9), i[3])
            # write project manager information
            ws.write('D{0}'.format(iteration+9), i[1])
            # write hours for each project per month
            ws.write_row('E{0}'.format(iteration+9), i[2], center)
            # write total hours field
            ws.write('{0}{1}'.format(total_col, iteration+9), sum(i[2]), center)
        except:
            print(k, i)

# Close indiv_plan_wkbook
indiv_plan_wkbook.close()






################################################################################
# Create Rollup Excel workbook
################################################################################
rollup_wkbook = xlsxwriter.Workbook(my_settings.out_rollup_file)
rollup_ws1 = rollup_wkbook.add_worksheet()

# Create formatting objects
big_bold = rollup_wkbook.add_format({'bold': 2, 'size': 22})
bold_1 = rollup_wkbook.add_format({'bold': 1})
bold_bkg_red = rollup_wkbook.add_format({'bold': 1, 'bg_color': '#F5A9BC', 'align': 'center'})
bold_bkg_yellow = rollup_wkbook.add_format({'bold': 1, 'bg_color': '#F2F5A9', 'align': 'center'})
bold_bkg_green = rollup_wkbook.add_format({'bold': 1, 'bg_color': '#04B404', 'align': 'center'})
bkg_red = rollup_wkbook.add_format({'bg_color': '#F5A9BC'})
bkg_yellow = rollup_wkbook.add_format({'bg_color': '#F2F5A9'})
bkg_green = rollup_wkbook.add_format({'bg_color': '#04B404'})
border_1 = rollup_wkbook.add_format({'border': 1})
merge_format = rollup_wkbook.add_format({'border': 1, 'align': 'center', 'bold': 1})
border_gray = rollup_wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD'})
border_gray_right = rollup_wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'right'})
border_gray_center = rollup_wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'center'})
percent_format = rollup_wkbook.add_format({'num_format': '0%'})

# Set column widths
rollup_ws1.set_column('A:A', 28)
rollup_ws1.set_column('B:M', 25)

# Set header content based on design
month_hdr_list = implement_design(my_settings.design, month_hdr)
wkg_hours_hdr_list = implement_design(my_settings.design, self.wkg_hours_hdr)
time_span_hdr_list = implement_design(my_settings.design, self.time_span_hdr)

# Calculate available hours sum
avail_hours_sum = sum(wkg_hours_hdr_list)

# Set design specific criteria
if my_settings.design == 'full_year':
    # set totals and percent column names
    totals_column = 'N'
    percent_column = 'O'
    # define date range for design
    date_range = 'Dec 27, {0} - Dec 25, {1}'.format(my_settings.fiscal_year, my_settings.fiscal_year+1)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('K9:M9', 'Quarter 1 - FY{0}'.format(int(my_settings.fy)+1), merge_format)

elif my_settings.design == 'quarter_2_3_4':
    # set totals column name
    totals_column = 'K'
    percent_column = 'L'
    # define date range for design
    date_range = 'Dec 27, {0} - Sep 30, {1}'.format(my_settings.fiscal_year - 1, my_settings.fiscal_year)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)

elif my_settings.design == 'quarter_2_3':
    # set totals column name
    totals_column = 'H'
    percent_column = 'I'
    # define date range for design
    date_range = 'Dec 27, {0} - Jun 26, {1}'.format(my_settings.fiscal_year - 1, my_settings.fiscal_year)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)

elif my_settings.design == 'quarter_2':
    # set totals column name
    totals_column = 'E'
    percent_column = 'F'
    # define date range for design
    date_range = 'Dec 27, {0} - Mar 27, {1}'.format(my_settings.fiscal_year - 1, my_settings.fiscal_year)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)

elif my_settings.design == 'quarter_3_4_1':
    # set totals column name
    totals_column = 'K'
    percent_column = 'L'
    # define date range for design
    date_range = 'Mar 28, {0} - Dec 25, {1}'.format(my_settings.fiscal_year, my_settings.fiscal_year + 1)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)
elif my_settings.design == 'quarter_3_4':
    # set totals column name
    totals_column = 'H'
    percent_column = 'I'
    # define date range for design
    date_range = 'Mar 28, {0} - Sep 30, {0}'.format(my_settings.fiscal_year)
    # write merged headers
    rollup_ws1.merge_range('B9:D9', 'Quarter 3 - FY{0}'.format(my_settings.fy), merge_format)
    rollup_ws1.merge_range('E9:G9', 'Quarter 4 - FY{0}'.format(my_settings.fy), merge_format)

# Write generic sheet content
rollup_ws1.write('A1', 'Staff Planning', big_bold)
rollup_ws1.write('A3', 'Staff Rollup - Only includes projects that are > 50% funding probability', bold_1)
rollup_ws1.write_row('A4', ['Manager:', 'TGM'])
rollup_ws1.write('A6', 'Key:')
rollup_ws1.write('A7', 'Key Explanation:')
rollup_ws1.write('B6', 'Trouble', bold_bkg_red)
rollup_ws1.write('B7', '<= 50% Covered', bold_bkg_red)
rollup_ws1.write('C6', 'Watch', bold_bkg_yellow)
rollup_ws1.write('C7', '51-80% Covered', bold_bkg_yellow)
rollup_ws1.write('D6', 'No Worries', bold_bkg_green)
rollup_ws1.write('D7', '> 81% Covered', bold_bkg_green)
rollup_ws1.write('A9', 'Technical Group', border_1)
rollup_ws1.write('A10', '', border_gray)
rollup_ws1.write('A11', 'Wkg Hrs Available =', border_gray_right)
rollup_ws1.write('A12', 'Processing Month =', border_gray_right)

# Write headers based on design
rollup_ws1.write_row('B10', month_hdr_list, border_gray_center)
rollup_ws1.write_row('B11', wkg_hours_hdr_list, border_gray_center)
rollup_ws1.write_row('B12', time_span_hdr_list, border_gray_center)
rollup_ws1.write('{0}10'.format(totals_column), 'Total', border_gray_center)
rollup_ws1.write('{0}11'.format(totals_column), avail_hours_sum, border_gray_center)
rollup_ws1.write('{0}12'.format(totals_column), date_range, border_gray_center)
rollup_ws1.write('{0}10'.format(percent_column), '', border_gray_center)
rollup_ws1.write('{0}11'.format(percent_column), 'Percent Covered', border_gray_center)
rollup_ws1.write('{0}12'.format(percent_column), '', border_gray_center)

# Create sorted staff name list from rollup dictionary
staff_column_list = [ k for k in rollup_dict.iterkeys() ]

# Sort staff column list
staff_column_list.sort()

# Write staff to sheet
rollup_ws1.write_column('A13', staff_column_list)

# Write the hours for each staff member
start_idx = 13
for staff in staff_column_list:

    # set staff hours list based on design
    if my_settings.design == 'quarter_3_4':
        staff_hours_list = rollup_dict[staff][:6]
    elif my_settings.design == 'quarter_2_3':
        staff_hours_list = rollup_dict[staff][:6]
    elif my_settings.design == 'quarter_2_3_4':
        staff_hours_list = rollup_dict[staff][:9]
    elif my_settings.design == 'quarter_3_4_1':
        staff_hours_list = rollup_dict[staff][:9]
    elif my_settings.design == 'full_year':
        staff_hours_list = rollup_dict[staff][:12]

    # retrieve staff hours list for staff member based on design
    #staff_hours_list = rollup_dict[staff][:6]#implement_design(design, rollup_dict[staff][:6]) # code was altered to accomodate q_3_4 setting via [:6]

    # calculate total hours
    total_hours = sum(staff_hours_list)

    # calculate percent covered
    percent_covered = (float(total_hours) / avail_hours_sum)

    # write hours for the appropriate cell range
    rollup_ws1.write_row('B{0}'.format(start_idx), staff_hours_list)

    # write total hours
    rollup_ws1.write('{0}{1}'.format(totals_column, start_idx), total_hours)

    # write percent covered
    rollup_ws1.write('{0}{1}'.format(percent_column, start_idx), percent_covered, percent_format)

    # advance index
    start_idx += 1

# Apply conditional formatting based on percent hours covered
# set data write ranges
data_start_row_index = 12
data_end_row_index = ((len(staff_column_list)-1) + data_start_row_index)

# run for each months available working hours
for i in range(0, (len(wkg_hours_hdr_list) + 1), 1):

    try:
        # get 50% hour value for trouble conditional bound
        trouble_max_value = (wkg_hours_hdr_list[i] * 0.5)

        # get 80% hour value for watch conditional bound
        watch_max_value = (wkg_hours_hdr_list[i] * 0.8)
    except:
        trouble_max_value = (avail_hours_sum * 0.5)
        watch_max_value = (avail_hours_sum * 0.8)

    # set conditional cell ranges
    start_cell = xlsxwriter.utility.xl_rowcol_to_cell(data_start_row_index, i+1)
    end_cell = xlsxwriter.utility.xl_rowcol_to_cell(data_end_row_index, i+1)

    # apply trouble condition
    trouble_condition = rollup_ws1.conditional_format('{0}:{1}'.format(start_cell, end_cell),
                        {'type': 'cell',
                         'criteria': '<=',
                         'value': trouble_max_value,
                         'format': bkg_red
                         })

    # apply watch condition
    watch_condition = rollup_ws1.conditional_format('{0}:{1}'.format(start_cell, end_cell),
                        {'type': 'cell',
                         'criteria': 'between',
                         'minimum': (trouble_max_value + 0.01),
                         'maximum': watch_max_value,
                         'format': bkg_yellow
                         })

    # apply no worries condition
    no_worries_condition = rollup_ws1.conditional_format('{0}:{1}'.format(start_cell, end_cell),
                        {'type': 'cell',
                         'criteria': '>',
                         'value': watch_max_value,
                         'format': bkg_green
                         })

# Close rollup workbook
rollup_wkbook.close()

################################################################################
# Format data for charts
chart_data_list = []
for k, v in projects_dict.iteritems():
    # assign data variables
    project = k
    staff_number_per_project = len(v)
    hours_sum_per_project = get_nested_list_sum(v, 2)
    # append data variabels to list
    chart_data_list.append([project, staff_number_per_project, hours_sum_per_project])

# Sort chart data appropriately
staff_per_prj_data = sorted(chart_data_list, key=operator.itemgetter(1), reverse=True)
hours_per_prj_data = sorted(chart_data_list, key=operator.itemgetter(2), reverse=True)

# Create summary workbook
summary_workbook = xlsxwriter.Workbook(my_settings.out_summary_file)

# Creating formatting objects
big_bold = summary_workbook.add_format({'bold': 2, 'size': 22})
bold_1 = summary_workbook.add_format({'bold': 1})
center_format = summary_workbook.add_format({'align': 'center'})
url_format = summary_workbook.add_format({'font_color': 'blue', 'underline': 0})

# Set hover over information
hyperlink_tip = 'Click name to open project workbook.'

# Add worksheet and chartsheets to summary book
summary_ws1 = summary_workbook.add_worksheet('summary')
summary_ws2 = summary_workbook.add_chartsheet('staff_per_project_graph')
summary_ws3 = summary_workbook.add_chartsheet('hours_per_project_graph')
summary_ws4 = summary_workbook.add_worksheet('staff_per_project_data')
summary_ws5 = summary_workbook.add_worksheet('hours_per_project_data')

# Set paper sizes for chartsheets
summary_ws2.set_paper(8)
summary_ws3.set_paper(8)

# Create summary worksheet
# calculate metrics
all_staff = len(ind_dict.keys())
all_projects = len(projects_dict.keys())
# set column widths
summary_ws1.set_column('A:A', 25)
summary_ws1.set_column('B:B', 15)
# write sheet
summary_ws1.write('A1', 'Staff Planning', big_bold)
summary_ws1.write('A3', 'Labor Planning Summary', bold_1)
summary_ws1.write('A5', 'Manager:')
summary_ws1.write('B5', 'TGM', center_format)
summary_ws1.write('A6', 'Number of Staff:')
summary_ws1.write('B6', all_staff, center_format)
summary_ws1.write('A7', 'Number of Projects:')
summary_ws1.write('B7', all_projects, center_format)

# Create chart data worksheets
# set column widths
summary_ws4.set_column('A:A', 25)
summary_ws4.set_column('B:C', 15)
summary_ws5.set_column('A:A', 25)
summary_ws5.set_column('B:C', 15)
# set chart data headers
chart_data_headers = ['Projects', 'Number of Staff', 'Number of Hours']

# add data for staff per project
summary_ws4.write_row('A1', chart_data_headers, bold_1)
for idx, row in enumerate(staff_per_prj_data):
    # create hyperlink if link location exists for project id
    link_location = project_path_dict[row[0]]
    summary_ws4.write_url('A{0}'.format(idx + 2), link_location, url_format, row[0], hyperlink_tip)
    summary_ws4.write('B{0}'.format(idx + 2), row[1])
    summary_ws4.write('C{0}'.format(idx + 2), row[2])
# add data for hours per project
summary_ws5.write_row('A1', chart_data_headers, bold_1)
for idx, row in enumerate(hours_per_prj_data):
    # create hyperlink if link location exists for project id
    link_location = project_path_dict[row[0]]
    summary_ws5.write_url('A{0}'.format(idx + 2), link_location, url_format, row[0], hyperlink_tip)
    summary_ws5.write('B{0}'.format(idx + 2), row[1])
    summary_ws5.write('C{0}'.format(idx + 2), row[2])

# Instantiate staff per project chart
staff_per_prj_chart = summary_workbook.add_chart({'type': 'column'})

# Configure staff per project data series
staff_per_prj_chart.add_series({
    'name':         '=staff_per_project_data!$B$1',
    'categories':   '=staff_per_project_data!$A$2:$A${0}'.format(len(staff_per_prj_data)+1),
    'values':       '=staff_per_project_data!$B$2:$B${0}'.format(len(staff_per_prj_data)+1)
    })

# Set staff per project chart properties
staff_per_prj_chart.set_title({'name': 'Staff Number per Project'})
staff_per_prj_chart.set_y_axis({'name': 'Number of Staff'})

# Set staff per project chart style and size
staff_per_prj_chart.set_style(18)

# Add staff per project chart to chartsheet
summary_ws2.set_chart(staff_per_prj_chart)


# Instantiate hours per project chart
hours_per_prj_chart = summary_workbook.add_chart({'type': 'column'})

# Configure hours per project data series
hours_per_prj_chart.add_series({
    'name':         '=hours_per_project_data!$C$1',
    'categories':   '=hours_per_project_data!$A$2:$A${0}'.format(len(staff_per_prj_data)+1),
    'values':       '=hours_per_project_data!$C$2:$C${0}'.format(len(staff_per_prj_data)+1)
    })

# Set hours per project chart properties
hours_per_prj_chart.set_title({'name': 'Hours per Project'})
hours_per_prj_chart.set_y_axis({'name': 'Number of Hours'})

# Set hours per project chart style
hours_per_prj_chart.set_style(18)

# Add hours per project chart to chartsheet
summary_ws3.set_chart(hours_per_prj_chart)

# Close rollup workbook
summary_workbook.close()

################################################################################


# Outro
#print "\nFinished Processing."
eg.message_box("Finished Processing.", "Progress.", "EXIT")
