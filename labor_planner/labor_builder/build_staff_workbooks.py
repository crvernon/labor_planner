import os
import collections
import operator
import xlrd
import xlsxwriter


def make_dirs(dir_list):
    """Create a directory if it does not exist for each in list

    :param dir_list:

    """

    for i in dir_list:

        if not os.path.exists(i):
            os.makedirs(i)


def open_xlsx(in_file):
    """Open and return workbook object

    :param in_file:                 Full path with file name and extension to the input file.

    :return:                        XLRD workbook object

    """

    msg = "The following file is either open by another user or cannot be opened:  {}".format(in_file)

    try:
        return xlrd.open_workbook(in_file)

    except IOError:
        raise(msg)


def add_to_dict(d, key, value):
    """Add key value pair to dictionary if it does not exist"""

    if key in d:
        d[key].append(value)
    else:
        d[key] = [value]


def create_staff_dict(key, value, s_dict={}):

    # iterate through each project component
    for v in value:
        # unpack values
        sn, fy, fte, prob = v
        # create staff dict key
        k = '{0}|{1}|{2}'.format(sn, fy, prob)
        # add staff info to dict; if in, create sum FTE
        if k in s_dict:
            # get previous FTE
            prev_fte = s_dict[k]
            # sum to get new FTE
            new_fte = (prev_fte + fte)
            # replace old FTE with new FTE
            s_dict[k] = new_fte
        else:
            s_dict[k] = fte
    return s_dict


def staff_dict_to_list(d, out_list=[]):

    for k in d.keys():

        v = d[k]

        sn, fy, prob = k.split('|')
        out_list.append([sn, fy, v, prob])

    return out_list


def get_pm_names(in_dict):
    pm_list = []
    for k in in_dict.keys():
        pm = k.split('|')[0]
        pm_list.append(pm)
    # sort list and make unique
    return sorted(list(set(pm_list)))


def get_staff_names(in_dict):
    staff_list = []

    for k in in_dict.keys():

        v = in_dict[k]

        for i in v:
            sn = i[0]
            staff_list.append(sn)
    # sort list and make unique
    return sorted(list(set(staff_list)))


def combine_list_unique(list_1, list_2):
    return sorted(list(set(list_1 + list_2)))


def format_file_name(string):
    sr = string.replace(',', '_').replace(' ', '_').replace(')', '_').replace('(', '_')
    return sr.replace('___', '_').replace('__', '_')


def process_input_data(in_file, target_fy):
    # set blank dicts
    temp_dict = {}
    final_dict = {}

    # open workbook
    wkbook = open_xlsx(in_file)

    # get staffing worksheet
    s = wkbook.sheet_by_name('Sheet1')

    # get column values starting at row 1
    id_col = s.col_values(0)[1:]
    title_col = s.col_values(1)[1:]
    #    desc_col = s.col_values(3)[1:]
    pm_col = s.col_values(5)[1:]
    #    fy_col = s.col_values(7)[1:]
    staff_col = s.col_values(4)[1:]
    fte_col = s.col_values(2)[1:]
    adj_fte_col = s.col_values(3)[1:]

    # calculate funding probability by dividing fte/adj_fte
    funding_prob_col = map(operator.truediv, adj_fte_col, fte_col)

    # combine lists
    out_list = zip(id_col, title_col, pm_col, staff_col, adj_fte_col, funding_prob_col)

    # create dictionary from list
    for i in out_list:
        # unpack values
        proj_id, title, pm, staff_name, fte, prob = i

        # get records for target FY
        fy = target_fy

        # create unique key
        key = '{0}|{1}|{2}'.format(pm, proj_id, title)
        val = [staff_name, fy, fte, prob]

        # add to dict
        add_to_dict(temp_dict, key, val)

    # process project dict to get correct FTE for each staff
    for key in temp_dict.keys():

        value = temp_dict[key]

        staff_dict = {}
        rep_list = []
        # create staff name dict
        create_staff_dict(key, value, staff_dict)

        # create list from staff dict
        staff_dict_to_list(staff_dict, rep_list)

        # replace current entry in project dict with new values
        final_dict[key] = rep_list

    return final_dict


def set_formatting(wkbook):
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


def percent_to_hours(wkg_hours_list, staff_percent_yr):
    total_hours = float(sum(wkg_hours_list))
    act_staff_hrs = staff_percent_yr * total_hours
    return [round((i / total_hours) * act_staff_hrs) for i in wkg_hours_list]


def write_static_worksheet_content(ws, fmt, fy, wkg_hrs_list):
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

    ws.write('A12', 'Group Staff', fmt[3])
    current_fy = int(fy[-2:])
    next_fy = 'FY{0}'.format(current_fy + 1)
    ws.merge_range('B12:D12', 'Quarter 2 - {0}'.format(fy), fmt[8])
    ws.merge_range('E12:G12', 'Quarter 3 - {0}'.format(fy), fmt[8])
    ws.merge_range('H12:J12', 'Quarter 4 - {0}'.format(fy), fmt[8])
    ws.merge_range('K12:M12', 'Quarter 1 - {0}'.format(next_fy), fmt[8])
    ws.write('N12', '', fmt[8])
    ws.write('O12', '', fmt[8])

    mth_list = ['',
                'Jan-{0}'.format(current_fy),
                'Feb-{0}'.format(current_fy),
                'Mar-{0}'.format(current_fy),
                'Apr-{0}'.format(current_fy),
                'May-{0}'.format(current_fy),
                'Jun-{0}'.format(current_fy),
                'Jul-{0}'.format(current_fy),
                'Aug-{0}'.format(current_fy),
                'Sep-{0}'.format(current_fy),
                'Oct-{0}'.format(current_fy),
                'Nov-{0}'.format(current_fy),
                'Dec-{0}'.format(current_fy),
                'Total']
    ws.write_row('A13', mth_list, fmt[6])
    ws.merge_range('O13:O14', '% of Available Hours Covered', fmt[6])

    total_hours = sum(wkg_hrs_list)
    ws.write('A14', 'Wkg Hrs Available =', fmt[6])
    ws.write_row('B14', wkg_hrs_list, fmt[6])
    ws.write('N14', total_hours, fmt[6])

    span_list = ['Processing Month =',
                 'Dec 27-Jan 23',
                 'Jan 24-Feb 20',
                 'Feb 21-Mar 27',
                 'Mar 28-Apr 24',
                 'Apr 25-May 22',
                 'May 23-Jun 26',
                 'Jun 27-Jul 24',
                 'Jul 25-Aug 21',
                 'Aug 22-Sept 30',
                 'Oct 1-Oct 23',
                 'Oct 24-Nov 20',
                 'Nov 21-Dec 25',
                 '',
                 '']
    ws.write_row('A15', span_list, fmt[6])


def populate_staff_info(ws, staff_list, key, value, wkg_hrs_list, fmt):
    # functions
    def get_staff_with_hours(sdict, value, staff_list, wkg_hrs_list):
        # for each staff member in group
        for sn in staff_list:

            # for each staff member on project
            for v in value:

                # unpack values
                staff_name, fy, fte, prob = v

                # if staff member is on this project
                if sn == staff_name:
                    # calculate hours per month per staff
                    act_wkg_hrs = percent_to_hours(wkg_hrs_list, fte)
                    # total project hours for staff member
                    act_hrs_total = sum(act_wkg_hrs)
                    # total percent that staff member hours will be of FY
                    tot_hrs_percent = (act_hrs_total / sum(wkg_hrs_list))
                    # add val to dict
                    sdict[sn] = [act_wkg_hrs, act_hrs_total, tot_hrs_percent]

    def get_staff_without_hours(sdict, staff_list, wkg_hrs_list):
        for sn in staff_list:
            if sn in sdict:
                pass
            else:
                act_wkg_hrs = [''] * len(wkg_hrs_list)
                sdict[sn] = [act_wkg_hrs, '', '']

    def write_staff_rows(ordered_sdict, start_row, ws, fmt):
        index = 0
        for k in ordered_sdict.keys():

            v = ordered_sdict[k]

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
            #            ws.write('N{0}'.format(row), act_hrs_total, f)
            sum_range = 'B{0}:M{0}'.format(row)
            tot_form = '{=SUM(' + sum_range + ')}'
            tot_cell = 'N{0}'.format(row)
            ws.write_formula(tot_cell, tot_form, f)
            #            ws.write('O{0}'.format(row), tot_hrs_percent, fp)
            per_form = '{=' + tot_cell + '/ (N14)}'
            ws.write_formula('O{0}'.format(row), per_form, fp)

            # advance row
            index += 1

        # return final staff row number + 1
        return row + 1

    def write_totals_row(ws, totals_row, start_row, f):

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

    # start row
    start_row = 16

    # locals
    sdict = {}

    # add data to dict where staff have hours on project
    get_staff_with_hours(sdict, value, staff_list, wkg_hrs_list)

    # add data to dict where staff do not have hours on the project
    get_staff_without_hours(sdict, staff_list, wkg_hrs_list)

    # order dictionary
    ordered_sdict = collections.OrderedDict(sorted(sdict.items()))

    # write staff information to worksheet; get row for totals
    totals_row = write_staff_rows(ordered_sdict, start_row, ws, fmt)

    # write totals row
    write_totals_row(ws, totals_row, start_row, fmt[4])


def populate_project_info(ws, k, v, fmt):
    # unpack key
    pm, pid, pname = k.split('|')
    # unpack val
    prob = float(v[0][3]) * 100
    # write project data
    ws.write('B8', prob, fmt[2])
    ws.write('B3', pid, fmt[2])
    ws.write('B4', pname, fmt[2])
    ws.write('B9', pm, fmt[2])


def create_project_worksheet(d, pm_name, wbook, staff_list, fmt, target_fy, wkg_hrs_list):

    for k in d.keys():

        v = d[k]

        if pm_name in k:
            # unpack vars
            pm, pid, pname = k.split('|')

            # create workbook for each project
            ws = wbook.add_worksheet(pid)

            # write static worksheet content
            write_static_worksheet_content(ws, fmt, target_fy, wkg_hrs_list)

            # write dynamic content
            populate_project_info(ws, k, v, fmt)
            populate_staff_info(ws, staff_list, k, v, wkg_hrs_list, fmt)


def create_empty_worksheets(wbook, staff_list, fmt, target_fy, wkg_hrs_list, num_sheets):
    # create the specified number of blank project sheets
    for i in range(1, (num_sheets + 1), 1):

        # set sheet name
        s = 'new_project_{0}'.format(i)

        # make blank worksheet
        ws = wbook.add_worksheet(s)

        # write static content
        write_static_worksheet_content(ws, fmt, target_fy, wkg_hrs_list)

        # write project cells
        ws.write('B8', '', fmt[2])
        ws.write('B3', '', fmt[2])
        ws.write('B4', '', fmt[2])
        ws.write('B9', '', fmt[2])

        # write staff area
        populate_staff_info(ws, staff_list, '', '', wkg_hrs_list, fmt)


def create_out_staff_file(staff_list, out_staff_file):
    # create out string of staff names delimited by semicolon
    out_string = ''
    for i in staff_list:
        out_string += '{0};'.format(i)

    # write out file
    with open(out_staff_file, 'w') as out:
        out.write(out_string)


if __name__ == '__main__':

    # vars
    in_dir = "/users/d3y010/projects/organizational/labor_planning/dale"
    in_file = "{0}/ecology_staffing.xlsx".format(in_dir)
    target_fy = 'FY18'
    out_sheets_dir = "{0}/FY_20{1}".format(in_dir, target_fy[-2:])
    admin_dir = "{0}/admin".format(in_dir)
    out_staff_file = "{0}/staff_list.csv".format(admin_dir)
    wkg_hrs_list = [152, 160, 200, 160, 160, 192, 152, 200, 192, 136, 120, 136]
    num_blank_wksheets = 10

    # make directories if they do not exits
    make_dirs([out_sheets_dir, admin_dir])

    # bring in data and output processed data as a dict
    data_dict = process_input_data(in_file, target_fy)

    # get a unique list of staff names
    map_staff_list = get_staff_names(data_dict)

    # get a unique list of pm names
    pm_list = get_pm_names(data_dict)

    # get a unique list of PMs and Staff
    staff_list = combine_list_unique(map_staff_list, pm_list)

    # write staff list output file
    create_out_staff_file(staff_list, out_staff_file)

    for index, pm_name in enumerate(staff_list):
        # format staff name to be file name
        file_name = format_file_name(pm_name.lower())

        # set out_file name
        wb_file = "{0}/{1}.xlsx".format(out_sheets_dir, file_name)

        # create Excel workbook for staff memmber
        wbook = xlsxwriter.Workbook(wb_file)

        # set workbook formatting
        fmt = set_formatting(wbook)

        # create project worksheet
        create_project_worksheet(data_dict, pm_name, wbook, staff_list, fmt, target_fy, wkg_hrs_list)

        # create blank worksheets
        create_empty_worksheets(wbook, staff_list, fmt, target_fy, wkg_hrs_list, num_blank_wksheets)

        # close workbook
        wbook.close()
