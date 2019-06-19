"""summary_wkbook.py

Build summary workbook with charts.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""

import operator
import xlsxwriter


class Summary:

    def __init__(self, config_obj, read_obj, data_obj):

        self.config_obj = config_obj

        self.data = data_obj

        self.read_obj = read_obj

        # Format data for charts
        chart_data_list = []
        for k in self.read_obj.projects_dict.keys():

            v = self.read_obj.projects_dict[k]

            project = k
            staff_number_per_project = len(v)
            hours_sum_per_project = sum([sum(i[2]) for i in v])

            # append data variables to list
            chart_data_list.append([project, staff_number_per_project, hours_sum_per_project])

        # Sort chart data appropriately
        staff_per_prj_data = sorted(chart_data_list, key=operator.itemgetter(1), reverse=True)
        hours_per_prj_data = sorted(chart_data_list, key=operator.itemgetter(2), reverse=True)

        # Create summary workbook
        summary_workbook = xlsxwriter.Workbook(self.config_obj.out_summary_file)

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

        # Set paper sizes for chart sheets
        summary_ws2.set_paper(8)
        summary_ws3.set_paper(8)

        # calculate metrics
        all_staff = len(self.read_obj.ind_dict.keys())
        all_projects = len(self.read_obj.projects_dict.keys())

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
            link_location = self.data.project_path_dict[row[0]]
            summary_ws4.write_url('A{0}'.format(idx + 2), link_location, url_format, row[0], hyperlink_tip)
            summary_ws4.write('B{0}'.format(idx + 2), row[1])
            summary_ws4.write('C{0}'.format(idx + 2), row[2])

        # add data for hours per project
        summary_ws5.write_row('A1', chart_data_headers, bold_1)

        for idx, row in enumerate(hours_per_prj_data):

            # create hyperlink if link location exists for project id
            link_location = self.data.project_path_dict[row[0]]
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

        # Add hours per project chart to chart sheet
        summary_ws3.set_chart(hours_per_prj_chart)

        # Close rollup workbook
        summary_workbook.close()
