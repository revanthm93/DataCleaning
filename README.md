# DataCleaning

1. unmergeexcelcells.py

Unmerge Excel Cells in Python and Populate the Top Cell Value.
In my current role, I had to load the data from excel sheet into SQL database. The excel sheets are vendor registration details and they have many merged cells, so I wrote a utility function in python using openpyxl library to unmerge the cell and populate the top cell value.

The first helper function create_merged_cell_lookup is taking the one worksheet and return a dict of cell location and its cell value.
Then this helper function will be passed into unmerge_cell_copy_top_value and we use the unmerge_cells from openpyxl then populate the value from create_merged_cell_lookup.
