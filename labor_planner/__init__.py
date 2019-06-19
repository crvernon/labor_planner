"""labor_planner

This package will generate labor planning Excel spreadsheets that provide overview,
project-level, staff labor projection summary, individual staff hours per project,
and summary charts from input staff planning spreadsheets.

@author Chris R. Vernon (chris.vernon@pnnl.gov)

"""

import pkg_resources

__version__ = pkg_resources.get_distribution('labor_planner').version

__all__ = ['main']