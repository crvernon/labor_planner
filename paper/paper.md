---
title: 'labor_planner: A Python package for planning and visualizing staff labor allocation'
tags:
  - Python
  - productivity
  - workforce
  - organization
  - management
authors:
  - name: Chris R. Vernon
    orcid: 0000-0002-3406-6214
    affiliation: 1
  - name: David R. Geist
    orcid: 0000-0002-4969-730X
    affiliation: 1
affiliations:
 - name: Pacific Northwest National Laboratory, Richland, WA, USA
   index: 1
date: 19 June 2019
bibliography: paper.bib
---

# Summary

Labor planning, a.k.a. workforce planning, has an established importance
in larger organizations due to the benefits of being prepared and aware of staff
allocation gaps that limit productivity and hamper a healthy work-life
balance [@Louch:2014]. @Blain:2019 presents two main elements of strategic
workforce planning:  1) understanding the current distribution of employee
skills and 2) trying to understand what will happen for staff in the future.
Additionally, the organizational benefits of providing staff with opportunity
conduct a healthy work-life balance are widely published (Examples...).

We developed the `labor_planner` Python package to provide insight into whether
staff are currently over- or under-committed and who, and which projects, are
the key consumers of staff chargeable hours.  Our package also evaluates and
visualizes projected labor allocation throughout a fiscal year to ensure staff
have time to complete existing project work and to assist staff who may not be
projected to meet billable requirements.  The `labor_planner` information for each
staff member are provided by project managers who distribute work among their
employees.  We find that project managers often do not communicate with one
another about staff time commitments when they create workforce plans for a
fiscal year.  `labor_planner` was created to ameliorate this unintended
miscommunication which is a common cause of limited productivity and resulting
poor work-life balance among staff.

`labor_planner` was created to be used by project managers, administrators, or
staff that wish to evaluate current and future labor allocation to increase the
efficacy of their projects. This package produces outputs in spreadsheet form
that are easily navigable and informative.  Outputs include staff-level
individual summaries detailing project commitment monthly, project-level
summaries highlighting the number of staff and hours per project within an
organization, staff overview summary containing funding probability
considerations and visualization, all staff rundown of monthly commitment tied
to visual indicators for over- and under-commitment, and overall summary charts
and data for all staff and all project and their interrelation.  `labor_planner`
was designed for extension and reuse and the authors encourage continued
community development.


# Acknowledgements

We acknowledge funding by the U.S. Department of Energy's Pacific Northwest
National Laboratory in Richland, Washington, United States of America as
managed by Battelle.

# References
