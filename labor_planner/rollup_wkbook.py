"""rollup_wkbook.py

Build rollup workbook for staff hours indicators of performance.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""

import xlsxwriter

import labor_planner.workbook_utils as util


class Rollup:

    def __init__(self, config_obj, read_obj, data_obj):

        self.config_obj = config_obj

        self.data = data_obj

        self.read_obj = read_obj

        rollup_wkbook = xlsxwriter.Workbook(self.config_obj.out_rollup_file)
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
        month_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.month_list)
        wkg_hours_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.wkg_hours_hdr)
        time_span_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.time_span_hdr)

        # Calculate available hours sum
        avail_hours_sum = sum(wkg_hours_hdr_list)

        # Set design specific criteria
        if self.config_obj.design == 'full_year':

            # set totals and percent column names
            totals_column = 'N'
            percent_column = 'O'

            # define date range for design
            date_range = 'Dec 27, {0} - Dec 25, {1}'.format(self.config_obj.fiscal_year, self.config_obj.fiscal_year+1)

            # write merged headers
            rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('K9:M9', 'Quarter 1 - FY{0}'.format(int(self.config_obj.fy)+1), merge_format)

        elif self.config_obj.design == 'quarter_2_3_4':

            # set totals column name
            totals_column = 'K'
            percent_column = 'L'

            # define date range for design
            date_range = 'Dec 27, {0} - Sep 30, {1}'.format(self.config_obj.fiscal_year - 1, self.config_obj.fiscal_year)

            # write merged headers
            rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)

        elif self.config_obj.design == 'quarter_2_3':

            # set totals column name
            totals_column = 'H'
            percent_column = 'I'

            # define date range for design
            date_range = 'Dec 27, {0} - Jun 26, {1}'.format(self.config_obj.fiscal_year - 1, self.config_obj.fiscal_year)

            # write merged headers
            rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)

        elif self.config_obj.design == 'quarter_2':

            # set totals column name
            totals_column = 'E'
            percent_column = 'F'

            # define date range for design
            date_range = 'Dec 27, {0} - Mar 27, {1}'.format(self.config_obj.fiscal_year - 1, self.config_obj.fiscal_year)

            # write merged headers
            rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)

        elif self.config_obj.design == 'quarter_3_4_1':

            # set totals column name
            totals_column = 'K'
            percent_column = 'L'

            # define date range for design
            date_range = 'Mar 28, {0} - Dec 25, {1}'.format(self.config_obj.fiscal_year, self.config_obj.fiscal_year + 1)

            # write merged headers
            rollup_ws1.merge_range('B9:D9', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('E9:G9', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('H9:J9', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)

        elif self.config_obj.design == 'quarter_3_4':

            # set totals column name
            totals_column = 'H'
            percent_column = 'I'

            # define date range for design
            date_range = 'Mar 28, {0} - Sep 30, {0}'.format(self.config_obj.fiscal_year)

            # write merged self.config_obj
            rollup_ws1.merge_range('B9:D9', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
            rollup_ws1.merge_range('E9:G9', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)

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
        staff_column_list = [k for k in self.read_obj.rollup_dict.keys()]

        # Sort staff column list
        staff_column_list.sort()

        # Write staff to sheet
        rollup_ws1.write_column('A13', staff_column_list)

        # Write the hours for each staff member
        start_idx = 13
        for staff in staff_column_list:

            # set staff hours list based on design
            if self.config_obj.design == 'quarter_3_4':
                staff_hours_list = self.read_obj.rollup_dict[staff][:6]

            elif self.config_obj.design == 'quarter_2_3':
                staff_hours_list = self.read_obj.rollup_dict[staff][:6]

            elif self.config_obj.design == 'quarter_2_3_4':
                staff_hours_list = self.read_obj.rollup_dict[staff][:9]

            elif self.config_obj.design == 'quarter_3_4_1':
                staff_hours_list = self.read_obj.rollup_dict[staff][:9]

            elif self.config_obj.design == 'full_year':
                staff_hours_list = self.read_obj.rollup_dict[staff][:12]

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
            # TODO:  raise specific exception
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
