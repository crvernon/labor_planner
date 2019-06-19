import xlsxwriter


class IndividualHours:

    def __init__(self, config_obj, read_obj, data_obj):

        self.config_obj = config_obj

        self.data = data_obj

        self.read_obj = read_obj

        indiv_plan_wkbook = xlsxwriter.Workbook(self.config_obj.out_individ_file)

        # Create formatting objects
        right_bold = indiv_plan_wkbook.add_format({'bold': 1, 'align': 'right'})
        bold_1 = indiv_plan_wkbook.add_format({'bold': 1})
        big_bold = indiv_plan_wkbook.add_format({'bold': 2, 'size': 22})
        merge_format = indiv_plan_wkbook.add_format({'border': 1, 'align': 'center'})
        center = indiv_plan_wkbook.add_format({'align': 'center'})
        right = indiv_plan_wkbook.add_format({'align': 'right'})
        url_format = indiv_plan_wkbook.add_format({'font_color': 'blue', 'underline': 0})

        # Set hover over information
        hyperlink_tip = 'Click name to open project workbook.'

        # Iterate through the individual hours dictionary and write content to file
        for k in self.read_obj.ind_dict.keys():

            v = self.read_obj.ind_dict[k]

            # format the worksheet name using modified staff name
            st_name = k.replace('*', '').strip()
            st_nospc = st_name.replace(' ', '_').replace('(', '').replace(')', '')
            ws_name = st_nospc.replace(',', '')

            # instantiate worksheet
            ws = indiv_plan_wkbook.add_worksheet(ws_name)

            # set column E and total column by design
            if self.config_obj.design == 'quarter_3_4':
                row_13_l = self.read_obj.month_header_with_total[3:9]
                row_13_l.append(self.read_obj.month_header_with_total[-1])
                row_14_l = self.read_obj.work_hours_row[3:9]
                row_14_l.append(sum(row_14_l))
                row_15_l = self.read_obj.total_month_range_row[3:9]
                row_15_l.append(self.read_obj.total_month_range_row[-1])
                end_col = 'J'
                total_col = 'K'

            elif self.config_obj.design == 'quarter_3_4_1':
                row_13_l = self.read_obj.month_header_with_total[-10:]
                row_14_l = self.read_obj.work_hours_row[3:12]
                row_14_l.append(sum(row_14_l))
                row_15_l = self.read_obj.total_month_range_row[-10:]
                end_col = 'M'
                total_col = 'N'

            elif self.config_obj.design == 'full_year':
                row_13_l = self.read_obj.month_header_with_total[-1:]
                row_14_l = self.read_obj.work_hours_row[0:12]
                row_14_l.append(sum(row_14_l))
                row_15_l = self.read_obj.total_month_range_row[0:12]
                end_col = 'P'
                total_col = 'Q'

            elif self.config_obj.design == 'quarter_2_3':
                row_13_l = self.read_obj.month_header_with_total[:6]
                row_13_l.append(self.data.month_header_with_total[-1])
                row_14_l = self.read_obj.work_hours_row[:6]
                row_14_l.append(sum(row_14_l))
                row_15_l = self.read_obj.total_month_range_row[:6]
                row_15_l.append(self.read_obj.total_month_range_row[-1])
                end_col = 'J'
                total_col = 'K'

            elif self.config_obj.design == 'quarter_2_3_4':
                row_13_l = self.read_obj.month_header_with_total[:9]
                row_14_l = self.read_obj.work_hours_row[:9]
                row_14_l.append(sum(row_14_l))
                row_15_l = self.read_obj.total_month_range_row[:9]
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
            if self.config_obj.design == 'quarter_3_4':
                ws.merge_range('E5:G5', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('H5:J5', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)
            elif self.config_obj.design == 'quarter_2_3':
                ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
            elif self.config_obj.design == 'quarter_3_4_1':
                ws.merge_range('E5:G5', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('H5:J5', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('K5:M5', 'Quarter 1 - FY{0}'.format(int(self.config_obj.fy) + 1), merge_format)
            elif self.config_obj.design == 'quarter_2_3_4':
                ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('K5:M5', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)
            elif self.config_obj.design == 'full_year':
                ws.merge_range('E5:G5', 'Quarter 2 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('H5:J5', 'Quarter 3 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('K5:M5', 'Quarter 4 - FY{0}'.format(self.config_obj.fy), merge_format)
                ws.merge_range('N5:P5', 'Quarter 1 - FY{0}'.format(int(self.config_obj.fy) + 1), merge_format)

            # ws.merge_range('L5:N5', 'Quarter 1 - FY17', merge_format)

            # write content to worksheet
            for iteration, i in enumerate(v):

                # create hyperlink if link location exists for project id
                link_location = self.data.project_path_dict[i[0]]
                ws.write_url('A{0}'.format(iteration + 9), link_location, url_format, i[0], hyperlink_tip)

                # write project funding probability
                ws.write('B{0}'.format(iteration + 9), i[4])

                # write project name
                ws.write('C{0}'.format(iteration + 9), i[3])

                # write project manager information
                ws.write('D{0}'.format(iteration + 9), i[1])

                # write hours for each project per month
                ws.write_row('E{0}'.format(iteration + 9), i[2], center)

                # write total hours field
                ws.write('{0}{1}'.format(total_col, iteration + 9), sum(i[2]), center)

        # Close indiv_plan_wkbook
        indiv_plan_wkbook.close()
