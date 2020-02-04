"""build_staff_workbooks.py

Build staff workbooks to hold hours data.

@author Chris R. Vernon (chris.vernon@pnnl.gov)
@license BSD 2-Clause

"""

import os
import collections

import numpy as np
import pandas as pd
import xlrd
import xlsxwriter


class BuildStaffWorkbooks:
    """Build staff workbook templates.  Each workbook will be used by the named staff member to forecast their,
     and those whom they manage, work hours for the calendar year.

     :param config_obj:                YAML configuration object

    """
    GROUP_STAFF_COL = "A"
    START_Q2_COL = "B"
    END_Q2_COL = "D"
    START_Q3_COL = "E"
    END_Q3_COL = "G"
    START_Q4_COL = "H"
    END_Q4_COL = "J"
    START_Q1_COL = "K"
    END_Q1_COL = "M"
    TOTAL_COL = "N"
    PER_COL = "O"

    def __init__(self, config_obj):

        self.staffing_file = config_obj.in_staff_csv
        self.target_fy = config_obj.fy
        self.out_staff_sheets_dir = config_obj.data_dir
        self.num_blank_wksheets = config_obj.num_blank_wksheets
        self.working_hours_file = config_obj.in_work_hours
        self.wkg_hrs_list, self.mth_list, self.mth_span_list = self.read_wkg_hrs()

        # build workbooks
        self.build()

    def read_wkg_hrs(self):
        """Process working hours file and format data as needed.

        :return:            [0] list of working hours
                            [1] list of months
                            [2] list of month span strings

        """
        df = pd.read_csv(self.working_hours_file)

        df['span'] = df['start_mon'] + ' ' + df['start_day'].astype(str) + '-' + df['end_mon'] + ' ' + df['end_day'].astype(str)

        wkg_hrs_list = df['work_hrs'].tolist()

        mth_list = df['month'].tolist()

        mth_span_list = df['span'].tolist()

        return wkg_hrs_list, mth_list, mth_span_list

    @staticmethod
    def open_xlsx(in_file):
        """Open and return workbook object

        :param in_file:                 Full path with file name and extension to the input file.

        :return:                        XLRD workbook object

        """

        msg = "The following file is either open by another user or cannot be opened:  {}".format(in_file)

        try:
            return xlrd.open_workbook(in_file)

        except IOError:
            raise msg

    @staticmethod
    def format_file_name(s):
        """Format the input string to accomodate the desired file name output.

        :param s:                       Input string

        :return:                        Formatted string

        """

        sr = s.replace(',', '_').replace(' ', '_').replace(')', '_').replace('(', '_')

        return sr.replace('___', '_').replace('__', '_').lower()

    @staticmethod
    def set_formatting(wkbook):
        """Set workbook formatting for cells.

        :param wkbook:                  Workbook object

        :return:                        Options for formatting

        """

        # formatting options
        bold = wkbook.add_format({'bold': 1})
        bold_big = wkbook.add_format({'bold': 2, 'size': 22})
        bkg_green = wkbook.add_format({'border': 1, 'bg_color': '#C1D4AB'})
        border_bold = wkbook.add_format({'border': 1, 'bold': 1})
        border_gray = wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD'})
        border_gray_right = wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'right'})
        border_gray_center = wkbook.add_format({'border': 1, 'bg_color': '#BDBDBD', 'align': 'center', 'text_wrap': 1})
        red_text_wrap = wkbook.add_format({'font_color': 'red', 'text_wrap': 1})
        bold_center_border = wkbook.add_format({'bold': 1, 'align': 'center', 'border': 1})
        num_percent = wkbook.add_format({'num_format': '0%', 'border': 1})
        bkg_blue = wkbook.add_format({'border': 1, 'bg_color': '#BED1DE'})
        border_1 = wkbook.add_format({'border': 1})
        num_percent_blue = wkbook.add_format({'num_format': '0%', 'border': 1, 'bg_color': '#BED1DE'})

        return bold, bold_big, bkg_green, border_bold, border_gray, border_gray_right, \
               border_gray_center, red_text_wrap, bold_center_border, num_percent, \
               bkg_blue, border_1, num_percent_blue

    def write_static_worksheet_content(self, ws, fmt, fy, wkg_hrs_list):
        """Write static (non-staff row) worksheet content

        :param ws:                      Worksheet object
        :param fmt:                     Formatting object
        :param fy:                      Two-digit fiscal year
        :param wkg_hrs_list:            List of working hours per month

        """

        # set column widths
        ws.set_column('A:A', 22)
        ws.set_column('B:N', 14)
        ws.set_column('O:O', 20)

        # write headers
        ws.write('A1', 'Project Labor Planning', fmt[1])

        ws.write('A3', '1a. Project Number:*', fmt[0])
        ws.write('B3', '', fmt[2])
        ws.write('D3', '1b. Proposal Number:*', fmt[0])
        ws.write('F3', '', fmt[2])
        ws.write('H3', '1c. WP#(s)/Task/WBS:*', fmt[0])
        ws.merge_range('J3:K3', '', fmt[2])

        ws.write('A4', '2. Title:', fmt[0])
        ws.merge_range('B4:I4', '', fmt[2])

        ws.write('A5', '3. Client:', fmt[0])
        ws.merge_range('B5:I5', '', fmt[2])

        ws.write('A6', '4. Start & End Dates:', fmt[0])
        ws.write('B6', '', fmt[2])
        ws.write('C6', 'through', fmt[0])
        ws.write('D6', '', fmt[2])

        ws.write('A7', '5. Funding amount:', fmt[0])
        ws.write('B7', '', fmt[2])

        # probability cell to be written in iterative portion of script
        ws.write('A8', '6. Probability %:', fmt[0])

        ws.write('A9', '7. Manager:', fmt[0])
        ws.merge_range('B9:D9', '', fmt[2])

        ws.write('A10', '8. Comments (optional):', fmt[0])
        ws.merge_range('B10:L10', '', fmt[2])

        p1 = "*Fill in either 1a. for a current project with funding; 1b. "
        p2 = "for a proposal that has not been funded/awarded yet; or 1c. for wp#(s), "
        p3 = "a task, or WBS that is funded from outside your group.  If this is a "
        p4 = "proposal (1b), please enter teh probability % of being funded/awarded "
        p5 = "in #6 above."
        text_fill = p1 + p2 + p3 + p4 + p5
        ws.merge_range('A11:N11', text_fill, fmt[7])

        ws.write('{}12'.format(BuildStaffWorkbooks.GROUP_STAFF_COL), 'Group Staff', fmt[3])
        current_fy = int(fy[-2:])
        next_fy = 'FY{0}'.format(current_fy + 1)

        ws.merge_range('{}12:{}12'.format(BuildStaffWorkbooks.START_Q2_COL, BuildStaffWorkbooks.END_Q2_COL),
                       'Quarter 2 - {0}'.format(fy), fmt[8])
        ws.merge_range('{}12:{}12'.format(BuildStaffWorkbooks.START_Q3_COL, BuildStaffWorkbooks.END_Q3_COL),
                       'Quarter 3 - {0}'.format(fy), fmt[8])
        ws.merge_range('{}12:{}12'.format(BuildStaffWorkbooks.START_Q4_COL, BuildStaffWorkbooks.END_Q4_COL),
                       'Quarter 4 - {0}'.format(fy), fmt[8])
        ws.merge_range('{}12:{}12'.format(BuildStaffWorkbooks.START_Q1_COL, BuildStaffWorkbooks.END_Q1_COL),
                       'Quarter 1 - {0}'.format(next_fy), fmt[8])

        ws.write('{}12'.format(BuildStaffWorkbooks.TOTAL_COL), '', fmt[8])
        ws.write('{}12'.format(BuildStaffWorkbooks.PER_COL), '', fmt[8])

        mth_list = [''] + ['{}-{}'.format(i, self.target_fy) for i in self.mth_list] + ['Total']

        ws.write_row('{}13'.format(BuildStaffWorkbooks.GROUP_STAFF_COL), mth_list, fmt[6])
        ws.merge_range('{0}13:{0}14'.format(BuildStaffWorkbooks.PER_COL), '% of Available Hours Covered', fmt[6])

        total_hours = sum(wkg_hrs_list)
        ws.write('{}14'.format(BuildStaffWorkbooks.GROUP_STAFF_COL), 'Wkg Hrs Available =', fmt[6])
        ws.write_row('{}14'.format(BuildStaffWorkbooks.START_Q2_COL), wkg_hrs_list, fmt[6])
        ws.write('{}14'.format(BuildStaffWorkbooks.TOTAL_COL), total_hours, fmt[6])

        span_list = ['Processing Month ='] + self.mth_span_list + ['', '']

        ws.write_row('{}15'.format(BuildStaffWorkbooks.GROUP_STAFF_COL), span_list, fmt[6])

    @staticmethod
    def create_staff_row_content(staff_list, wkg_hrs_list):
        """Create content to populate staff rows for each worksheet.

        :param staff_list:              List of all staff
        :param wkg_hrs_list:            List of working hours for each month

        :return:                        Dictionary of blank content for each staff member

        """

        d = {}

        for sn in staff_list:

            act_wkg_hrs = [''] * len(wkg_hrs_list)
            d[sn] = [act_wkg_hrs, '', '']

        return d

    @staticmethod
    def write_staff_rows(ordered_dict, start_row, ws, fmt):
        """"Write staff row content.

        :param ordered_dict:           Ordered dictionary of staff and their associated content
        :param start_row:               Row to start writing staff content on in the worksheet
        :param ws:                      Worksheet object
        :param fmt:                     Formatting object

        """

        for index, k in enumerate(ordered_dict.keys()):

            v = ordered_dict[k]

            # set row number
            row = start_row + index

            # set formatting
            if index % 2 == 0:
                f = fmt[11]
                fp = fmt[9]
            else:
                f = fmt[10]
                fp = fmt[12]

            # unpack values
            act_wkg_hrs, act_hrs_total, tot_hrs_percent = v

            # write info
            ws.write('A{0}'.format(row), k, f)
            ws.write_row('B{0}'.format(row), act_wkg_hrs, f)

            sum_range = 'B{0}:M{0}'.format(row)
            tot_form = '{=SUM(' + sum_range + ')}'
            tot_cell = 'N{0}'.format(row)
            ws.write_formula(tot_cell, tot_form, f)

            per_form = '{=' + tot_cell + '/ (N14)}'
            ws.write_formula('O{0}'.format(row), per_form, fp)

        # return final staff row number + 1
        return row + 1

    @staticmethod
    def write_totals_row(ws, totals_row, start_row, f):
        """Calculate the monthly totals rows and write them to the worksheet.

        :param ws:              Worksheet object
        :param totals_row:      Totals row number
        :param start_row:       Start row number
        :param  f:              Cell format

        """

        # set alpha string to be called by index
        alpha_str = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        ws.write('A{0}'.format(totals_row), 'Total', f)

        # from B to N
        for i in range(1, 14, 1):

            col = alpha_str[i]
            target_cell = '{0}{1}'.format(col, totals_row)
            end_row = totals_row - 1
            sum_range = '{0}{1}:{0}{2}'.format(col, start_row, end_row)
            formula = '{=SUM(' + sum_range + ')}'
            ws.write_formula(target_cell, formula, f)

        ws.write('O{0}'.format(totals_row), '', f)

    def populate_staff_info(self, ws, staff_list, wkg_hrs_list, fmt):
        """Create staff rows in worksheets.

        :param ws:                      Worksheet object
        :param staff_list:              List of all staff
        :param wkg_hrs_list:            List of working hours for each month
        :param fmt:                     Formatting object

        """

        # start row
        start_row = 16

        # add data to dict where staff do not have hours on the project
        d = self.create_staff_row_content(staff_list, wkg_hrs_list)

        # order dictionary
        ordered_dict = collections.OrderedDict(sorted(d.items()))

        # write staff information to worksheet; get row for totals
        totals_row = self.write_staff_rows(ordered_dict, start_row, ws, fmt)

        # write totals row
        self.write_totals_row(ws, totals_row, start_row, fmt[4])

    def create_empty_worksheets(self, wbook, staff_list, fmt, target_fy, wkg_hrs_list, num_sheets):
        """Create empty worksheets for each workbook

        :param wbook:                   Workbook object
        :param staff_list:              List of all staff
        :param fmt:                     Formatting object
        :param target_fy:               Two-digit target year
        :param wkg_hrs_list:            List of working hours per month
        :param num_sheets:              The number of template sheets to create

        """

        # create the specified number of blank project sheets
        for i in range(1, (num_sheets + 1), 1):

            # set sheet name
            s = 'new_project_{0}'.format(i)

            # make blank worksheet
            ws = wbook.add_worksheet(s)

            # write static content
            self.write_static_worksheet_content(ws, fmt, target_fy, wkg_hrs_list)

            # write project cells
            ws.write('B8', '', fmt[2])
            ws.write('B3', '', fmt[2])
            ws.write('B4', '', fmt[2])
            ws.write('B9', '', fmt[2])

            # write staff area
            self.populate_staff_info(ws, staff_list, wkg_hrs_list, fmt)

    def read_staff_file(self):
        """Read and process input staff file.  A labor planning workbook will be generated for each
        staff member in this file.

        :return:                            List of staff full names

        """
        df = pd.read_csv(self.staffing_file)

        df.fillna('', inplace=True)

        df['middle_initial'] = df['middle_initial'].replace(' ', '')

        df['full_name']= np.where(df['middle_initial'] == '',
                                     df['last_name'] + ', ' + df['first_name'],
                                     df['last_name'] + ', ' + df['first_name'])

        return df['full_name'].tolist()

    def build(self):
        """Method to build all staff worksheets."""

        # make staff sheet directory if it does not exist
        if not os.path.exists(self.out_staff_sheets_dir):
            os.makedirs(self.out_staff_sheets_dir)

        staff_list = self.read_staff_file()

        for pm_name in staff_list:

            # set out_file name
            wb_file = os.path.join(self.out_staff_sheets_dir, "{}.xlsx".format(self.format_file_name(pm_name)))

            with xlsxwriter.Workbook(wb_file) as wbook:

                # set workbook formatting
                fmt = self.set_formatting(wbook)

                # create blank worksheets
                self.create_empty_worksheets(wbook, staff_list, fmt, self.target_fy, self.wkg_hrs_list, self.num_blank_wksheets)
