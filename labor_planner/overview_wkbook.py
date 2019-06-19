import os
import xlsxwriter

import labor_planner.workbook_utils as util


class Overview:

    def __init__(self, config_obj, data_obj):

        self.config_obj = config_obj

        self.data = data_obj

        # Create graph Excel workbook and instantiate worksheet
        graph_workbook = xlsxwriter.Workbook(self.config_obj.out_overview_file)

        graph_ws1 = graph_workbook.add_worksheet()

        # Create formatting objects
        header_merge_format = util.enable_formatting(graph_workbook)[12]
        bold_1 = util.enable_formatting(graph_workbook)[2]
        bold_center = util.enable_formatting(graph_workbook)[3]
        url_format = util.enable_formatting(graph_workbook)[11]

        # List of headers
        merged_header_text = 'Proportion of FTE Covered in FY{0}'.format(self.config_obj.fy)
        header_list = ['> 50% Funding Probability', '<= 50% Funded Probability', 'All Projects']

        # Set hover over information
        hyperlink_tip = 'Click name to open source workbook.'

        # Set column widths
        graph_ws1.set_column('A:A', 20)
        graph_ws1.set_column('B:D', 28)

        # Write headers to worksheet
        graph_ws1.merge_range('B1:D1', 'Proportion of FTE Covered in FY{0}'.format(self.config_obj.fy), header_merge_format)
        graph_ws1.write('A2', 'Staff Member', bold_1)
        graph_ws1.write_row('B2', header_list, bold_center)

        # Write data to worksheet
        graph_ws1.write_column('A3', self.data.full_name_list)
        graph_ws1.write_column('B3', self.data.full_high_list)
        graph_ws1.write_column('C3', self.data.full_low_list)
        graph_ws1.write_column('D3', self.data.percent_list)

        # Create hyperlinks
        # --create hyperlink for each staff member with an existing workbook
        # --start at A3, idx = 3
        idx = 3
        for i in self.data.full_name_list:
            # set hyperlink parameters
            cell_location = "A{0}".format(idx)

            # format possible name issues
            l_name = i.replace('*','').strip()
            l_nospc = l_name.replace(' ','_').replace('(','').replace(')','')
            l = l_nospc.replace(',','')

            # create link location
            link_location = "{0}#{1}!a1".format(os.path.basename(self.config_obj.out_individ_file), l)

            # create hyperlink if link location exists
            graph_ws1.write_url(cell_location, link_location, url_format, i, hyperlink_tip)

            # advance index for cell location
            idx += 1

        # Create chart object
        chart = graph_workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

        # Add total percent covered series
        chart.add_series({
            'name':         'Prob > 50%',
            'categories':   '=Sheet1!$A$3:$A${0}'.format(self.data.end_row+1),
            'values':       '=Sheet1!$B$3:$B${0}'.format(self.data.end_row+1)
            })

        # Add probable (<= 50%) percent covered series
        chart.add_series({
            'name':         'Prob <= 50%',
            'categories':   '=Sheet1!$A$3:$A${0}'.format(self.data.end_row+1),
            'values':       '=Sheet1!$C$3:$C${0}'.format(self.data.end_row+1)
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
