import os


class Stage:

    def __init__(self, config_obj, read_obj):

        self.config_obj = config_obj

        # get information from staff workbooks
        self.read_obj = read_obj

        # header containing the working hours by design
        self.wkg_hours_hdr_list = self.implement_design(self.read_obj.wkg_hours_hdr)

        # get the total available hours for the design term
        self.avail_hours_sum = sum(self.wkg_hours_hdr_list)

        # list of post masters staff
        self.post_ma_list = []

        # list of all staff
        self.combine_list = []

        # combined list for formatted staff names
        self.name_list = []

        # list of percent covered
        self.percent_list = []

        # sum of post masters hours
        self.post_masters_hours = []

        # create list of high and low probability of funding
        self.high_prob_list = []
        self.low_prob_list = []

        # list of formatted staff names associated with high and low probability funding
        self.high_name_list = []
        self.low_name_list = []

        # list of high and low probability percent funded
        self.high_percent_list = []
        self.low_percent_list = []

        # split post grad list and combined list
        self.process_non_staff()

        # create high and low probability lists
        self.calc_high_probability()
        self.calc_low_probability()

        self.look_up_list = []
        self.full_name_list = []
        self.full_high_list = []
        self.full_low_list = []

        # Get value for the length of name column
        self.end_row = self.prep_chart_info()

        # add project id and path to sheet to dictionary
        self.project_path_dict = self.set_project_path_dict()

    def implement_design(self, in_list):
        """Gets a portion of a list that is based on the specified design

        :param in_list:             List of 12 values for each month in the year

        :return:                    Truncated list based on the design type

        """
        if self.config_obj.design == 'full_year':
            out_list = in_list

        elif self.config_obj.design == 'quarter_2_3_4':
            out_list = in_list[0:9]

        elif self.config_obj.design == 'quarter_2_3':
            out_list = in_list[0:6]

        elif self.config_obj.design == 'quarter_2':
            out_list = in_list[0:3]

        elif self.config_obj.design == 'quarter_3_4_1':
            out_list = in_list[-9:]

        elif self.config_obj.design == 'quarter_3_4':
            out_list = in_list[3:9]

        else:
            out_list = []

        return out_list

    def process_non_staff(self):
        """Create new list removing **Post MA assessments, keep them in a separate list"""

        for k in self.read_obj.staff_dict.keys():

            v = self.read_obj.staff_dict[k]

            if '**' in k:
                self.post_ma_list.append(sum(v))
            else:
                # calculate percent covered
                total_percent_covered = round((float(sum(v)) / float(self.avail_hours_sum)), 2)
                self.combine_list.append([k, total_percent_covered])

        # Sort by last name ascending
        self.combine_list.sort()

        # Create separate list for names and percent covered
        self.name_list = [i[0] for i in self.combine_list]
        self.percent_list = [i[1] for i in self.combine_list]

        # Return post masters list
        self.post_masters_hours = sum(self.post_ma_list)

    def calc_high_probability(self):
        """Create lists of high-probability funding."""

        for k in self.read_obj.staff_high_prob_dict.keys():

            v = self.read_obj.staff_high_prob_dict[k]

            if '**' in k:
                pass

            else:
                high_prob_percent = round((float(sum(v)) / float(self.avail_hours_sum)), 2)
                self.high_prob_list.append([k, high_prob_percent])

        # Sort by last name ascending
        self.high_prob_list.sort()

        # Create separate lists for names and percent covered
        self.high_name_list = [i[0] for i in self.high_prob_list]
        self.high_percent_list = [i[1] for i in self.high_prob_list]

    def calc_low_probability(self):
        """Create lists of low-probability funding."""

        for k in self.read_obj.staff_low_prob_dict.keys():

            v = self.read_obj.staff_low_prob_dict[k]

            if '**' in k:
                pass
            else:
                low_prob_percent = round((float(sum(v)) / float(self.avail_hours_sum)), 2)
                self.low_prob_list.append([k, low_prob_percent])

        # Sort by last name ascending
        self.low_prob_list.sort()

        # Create separte list for names and percent covered
        self.low_name_list = [i[0] for i in self.low_prob_list]
        self.low_percent_list = [i[1] for i in self.low_prob_list]

    def prep_chart_info(self):
        """Create lists that will be used to populate chart values

        :return:                    Value for the length of the name column

        """

        for i in self.name_list:

            if (i in self.high_name_list) and (i in self.low_name_list):
                self.look_up_list.append([i, self.high_percent_list[self.high_name_list.index(i)], self.low_percent_list[self.low_name_list.index(i)]])

            elif (i in self.high_name_list) and (i not in self.low_name_list):
                self.look_up_list.append([i, self.high_percent_list[self.high_name_list.index(i)], 0.0])

            elif (i in self.low_name_list) and (i not in self.high_name_list):
                self.look_up_list.append([i, 0.0, self.low_percent_list[self.low_name_list.index(i)]])

            else:
                self.look_up_list.append([i, 0.0, 0.0])

        # Sort look up list
        self.look_up_list.sort()

        # Create lists for outputs
        self.full_name_list = [i[0] for i in self.look_up_list]
        self.full_high_list = [i[1] for i in self.look_up_list]
        self.full_low_list = [i[2] for i in self.look_up_list]

        # Get value for the length of name column
        return len(self.full_name_list) + 1

    def set_project_path_dict(self):
        """Add project id and path to sheet in a dictionary.

        :returns:               Dictionary of project paths for workbook sheets
        """

        project_path_dict = {}

        for idx, k in enumerate(self.read_obj.projects_dict.keys()):

            ws = 'sheet_{}'.format(idx)

            if k not in project_path_dict:
                project_path_dict[k] = "{0}#{1}!a1".format(os.path.basename(self.config_obj.out_project_file), ws)

            else:
                raise RuntimeError("ERROR:  Duplicate project IDs detected for:  {}".format(k))

        return project_path_dict
