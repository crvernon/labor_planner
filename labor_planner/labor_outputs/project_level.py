"""project_level.py

Build project level workbook summary.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""


import xlsxwriter

import labor_planner.workbook_utils as util


class Projects:

    def __init__(self, config_obj, read_obj, data_obj):
        self.config_obj = config_obj

        self.data = data_obj

        self.read_obj = read_obj

        # Create projects Excel workbook and instantiate worksheet
        project_wkbook = xlsxwriter.Workbook(self.config_obj.out_project_file)

        # Create formatting objects
        bold_1 = util.enable_formatting(project_wkbook)[2]
        big_bold = util.enable_formatting(project_wkbook)[4]
        border_1 = util.enable_formatting(project_wkbook)[6]
        merge_format = util.enable_formatting(project_wkbook)[13]
        border_gray = util.enable_formatting(project_wkbook)[7]
        border_gray_right = util.enable_formatting(project_wkbook)[8]
        border_gray_center = util.enable_formatting(project_wkbook)[9]
        percent_format = util.enable_formatting(project_wkbook)[10]

        # Set worksheet index for naming convention
        ws_index = 0

        # Iterate through project dictionary
        for k in self.read_obj.projects_dict.keys():

            v = self.read_obj.projects_dict[k]

            # get project manager name
            pm = v[0][1]
            pb = v[0][3]

            # instantiate worksheet
            set_ws_name = 'sheet_{0}'.format(ws_index)
            prj_ws = project_wkbook.add_worksheet(set_ws_name)
            ws_index += 1

            # Set column widths
            prj_ws.set_column('A:A', 20)
            prj_ws.set_column('B:M', 24)

            # Set header content based on design
            month_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.month_list)
            wkg_hours_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.wkg_hours_hdr)
            time_span_hdr_list = util.implement_design(self.config_obj.design, self.read_obj.time_span_hdr)

            # Calculate available hours sum
            avail_hours_sum = sum(wkg_hours_hdr_list)

            # Set design specific criteria
            des_inst = util.set_merge_range(prj_ws, self.config_obj.design, self.config_obj.fiscal_year, 9, merge_format)
            totals_column = des_inst[0]
            percent_column = des_inst[1]
            date_range = des_inst[2]

            # Write generic sheet content
            prj_ws.write('A1', 'Staff Planning', big_bold)
            prj_ws.write('A3', 'Project Rollup', bold_1)
            prj_ws.write_row('A4', ['Project ID:', k])
            prj_ws.write_row('A5', ['Project Title:', self.read_obj.prj_title_dict[k]])
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