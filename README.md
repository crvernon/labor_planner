[![DOI](https://zenodo.org/badge/192757129.svg)](https://zenodo.org/badge/latestdoi/192757129)
[![Build Status](https://travis-ci.org/crvernon/labor_planner.svg?branch=master)](https://travis-ci.org/crvernon/labor_planner)

# labor_planner
Plan and visualize staff calendar year labor allocation

## Contact
Chris R. Vernon (chris.vernon@pnnl.gov)

## Getting Started
The `labor_planner` package uses only Python 3.3 and up.

### Step 1:
Clone the repository into your desired directory:

`git clone https://github.com/crvernon/labor_planner`

### Step 2:
You can install `labor_planner` by running the following from your cloned directory (NOTE: ensure that you are using the desired `python` instance):

`python setup.py install`

### Step 3:
Confirm that the module and its dependencies have been installed by running from your prompt:

```python
from labor_planner import LaborPlanner
```

If no error is returned then you are ready to go!

## Setting up a run

### Setup the `config.yml` file
There is an example config file in the `labor_planner/example` directory of this package that describes each input.

There are three main blocks using in the `labor_planner` configuration file:  `project`, `builder`, and `planner`.  Each block contains key, value pairs for each required setting.  The following details available settings:

#### `project` block:

| key | description |
| -- | -- |
| `input_directory` | "full path to the directory containing the staff spreadsheet directories" |
| `staff_file` | "full path with file name and extension to the input _all_staff.csv_ file" |
| `work_hours_csv` | "full path with file name and extension to the _work_hours.csv_ file" |
| `fiscal_year` | Four digit year as an integer. E.g., 2019 |
| `staff_workbook_dir` | "full path to the directory where the staff labor planning workbooks are stored" |
| `build_workbooks` | Boolean.  True or False.  `True` to build a staff workbook for each staff member listed in the all_staff.csv file.  NOTE:  Change the `staff_workbook_dir` path before running this or the example worksheets will be overwritten. Also, this can not be `True` if `run_labor_planner` is also set to `True`. |
| `run_labor_planner` | Boolean.  True or False.  `True` to run the labor planner and generate summary outputs.  NOTE, this can not be `True` if `build_workbooks` is also set to `True`. |

#### `builder` block:

| key | description |
| -- | -- |
| `num_blank_wksheets` | Integer value for the number of blank projects to generate |

#### `planner` block:

| key | description |
| -- | -- |
| `output_directory` | "full path with to the directory where the outputs will be written" |
| `run_design` | Either "full_year", "quarter_2_3", "quarter_3_4", or "quarter_3_4_1" |

### Setup the reference files
There are two reference files that are necessary to run this package (examples included in package):

- `all_staff.csv`:  This file is a comma separated file with three columns:  last_name, first_name, and middle_initial.  The following header must be present:  `last_name`, `first_name`, and `middle_initial`.

- `work_hours.csv`:  This file contains month abbreviation, start month, start day, end month, end day, and work hours associated with each month for the calendar year.  The following header must be present:  `month`, `start_mon`, `start_day`, `end_mon`, `end_day`, `work_hrs`.

### Setup the staff data files
An Excel workbook must exist for each staff member.  The name of the workbook must be in the following format: `<lastname>_<firstname>_<middleinitial-if-exists>.xlsx`.  See examples for formatted workbooks.  Each workbook has multiple worksheets.  Each worksheet is representative of one project that the staff member manages.  Only the managing staff member should fill out a sheet for their project.  This sheet should include the projected hours for all staff that are participating on their project.  All staff workbooks should be nested in a directory named like the following: `FY_<four-digit-calendar-year>`.  Template sheets for each staff member can be generated by setting up the _all_staff.csv_ file and then setting the YAML config file option to `True` for the `build_workbooks` option.  Options for `build_workbooks` and `run_labor_planner` cannot both be set to `True` at the same time.  This is due to the staff data files needed to be populated by staff before the planner generates outputs that have meaningful content.

## Running `labor_planner`

### Running from terminal or command line
Ensure that you are using the desired `python` instance then run:

`python <path-to-labor_planner-module>/main.py <path-to-the-config-file>`

### Running from a Python Prompt or from another script

```python
from labor_planner import LaborPlanner
LaborPlanner('<path-to-config-file>')
```

## Outputs
The following five outputs will be saved to the outputs directory assigned in the config file:
- `overview_chart.xlsx`:  Contains a single worksheet showing the probability of funding for each staff member with a link to their individual planning worksheet. Also provides a bar chart of funding per probability range.  Staff member names are linked to their corresponding `individual_staff_summary.xlsx` sheets.
- `projects.xlsx`:  Contains a worksheet for every project and the staff that contribute to them.
- `individual_staff_summary.xlsx`:  Contains a worksheet for each individual staff member that details each project, funding probability, project hours per month and a link to the associated project worksheet.  Each project number is linked to their corresponding `project.xlsx` sheet.
- `rollup.xlsx`:  Contains a single worksheet that highlights the degree of funding for each staff member per month.
- `summary.xlsx`:  Contains worksheets for total staff and projects; charts for staff per project, hours per project; data with links for staff per project and total hours.  Tabular summary worksheets allow a linkage between the project number and the associated `projects.xlsx` sheet.

## Community involvement
`labor_planner` was built to be extensible.  It is our hope that the community will continue the development of this software.  Please submit a pull request for any work that you would like have considered as a core part of this package.  You will be properly credited for your work and it will be distributed under our current open-source license.  Any issues should be submitted through standard GitHub issue protocol and I will deal with these promptly.  
