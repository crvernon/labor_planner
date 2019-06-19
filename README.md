# labor_planner
Plan and visualize staff fiscal year labor allocation

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
Confirm that the module and its dependencies have been installed by running from your Prompt:

```python
from labor_planner import LaborPlanner
```

If no error is returned then you are ready to go!

## Setting up a run

### Setup the `config.yaml` file
There is an example config file in the `labor_planner/example` directory of this package that describes each input.

### Setup the reference files
There are two reference files that are necessary to run this package (examples included in package):
- `staff_list.csv`:  This file is a semicolon separated list of staff names that match the names of each individual in the input Excel staff sheets
- `work_hours.csv`:  This file contains month abbreviation, start month, start day, end month, end day, and work hours associated with each month for the fiscal year

### Setup the staff data files
An Excel workbook must exist for each staff member.  The name of the workbook must be in the following format: `<lastname>_<firstname>_<middleinitial-if-exists>.xlsx`.  See examples for formatted workbooks.  Each workbook has multiple worksheets.  Each worksheet is representative of one project that the staff member manages.  Only the managing staff member should fill out a sheet for their project.  This sheet should include the projected hours for all staff that are participating on their project.  All staff workbooks should be nested in a directory named like the following: `FY_<four-digit-fiscal-year`.

## Running `labor_planner`

### Running from terminal or command line
Ensure that you are using the desired `python` instance then run:

`python <path-to-labor_planner-module>/main.py <path-to-the-config-file>`

### Running from a Python Prompt

```python
from labor_planner import LaborPlanner
LaborPlanner('<path-to-config-file')
```

### Running from another Python script:

```python
from labor_planner import LaborPlanner
LaborPlanner('<path-to-config-file')
```

## Outputs
The following five outputs will be saved to the outputs directory assigned in the config file:
- `overview_chart.xlsx`:  Contains a single worksheet showing the probability of funding for each staff member with a link to their individual planning worksheet. Also provides a bar chart of funding per probability range.
- `projects.xlsx`:  Contains a worksheet for every project and the staff that contribute to them.
- `individual_staff_summary.xlsx`:  Contains a worksheet for each individual staff member that details each project, funding probability, project hours per month and a link to the associated project worksheet.
- `rollup.xlsx`:  Contains a single worksheet that highlights the degree of funding for each staff member per month.
- `summary.xlsx`:  Contains worksheets for total staff and projects; charts for staff per project, hours per project; data with links for staff per project and total hours.
