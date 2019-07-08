"""workbook_utils.py

Broadly used workbook utility functions.

@author Chris R. Vernon (chris.vernon@pnnl.gov)
@license BSD 2-Clause

"""


def enable_formatting(in_workbook_instance):
    """Index options for preconfigured formatting:

    :param in_workbook_instance:                xlsxwriter workbook instance

    :returns:
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

    return align_center, align_right, bold_1, bold_1_center, bold_big, \
            bold_right, border_1, border_gray, border_gray_right, \
            border_gray_center, num_percent, url, merge_header, merge_border, \
            bkg_red, bkg_yellow, bkg_green, bkg_red_bold, bkg_yellow_bold, \
            bkg_green_bold


def implement_design(design, in_list):
    """Gets a portion of a list that is based on the specified design

    :param in_list:             List of 12 values for each month in the year

    :return:                    Truncated list based on the design type

    """
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

    else:
        out_list = []

    return out_list


def set_merge_range(worksheet_instance, design, fiscal_year, start_row_num, format_type):
    """Create merged format strings for output workbook columns by design.

    """

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
