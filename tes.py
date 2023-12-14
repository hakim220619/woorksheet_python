import xlsxwriter

# Create a new workbook and add some worksheets
workbook = xlsxwriter.Workbook("outline.xlsx")
worksheet1 = workbook.add_worksheet("Outlined Rows")
worksheet2 = workbook.add_worksheet("Collapsed Rows")
worksheet3 = workbook.add_worksheet("Outline Columns")
worksheet4 = workbook.add_worksheet("Outline levels")

# Add a general format
bold = workbook.add_format({"bold": 1, 'border': True})


###############################################################################
#
# Example 1: A worksheet with outlined rows. It also includes SUBTOTAL()
# functions so that it looks like the type of automatic outlines that are
# generated when you use the Excel Data->SubTotals menu item.
#
# For outlines the important parameters are 'level' and 'hidden'. Rows with
# the same 'level' are grouped together. The group will be collapsed if
# 'hidden' is enabled. The parameters 'height' and 'cell_format' are assigned
# default values if they are None.
#


# Adjust the column width for clarity
worksheet1.set_column("A:A", 20)
worksheet1.set_column("B:B", 40)

worksheet1.write("B5", "A")
worksheet1.write("B6", "1")
# Add the data, labels and formulas
worksheet1.write("C5", "dfghsindibdycb vduyad", bold)
worksheet1.write("C6", "dfgsdfy")
worksheet1.write("C7", "North")
worksheet1.write("C8", "North")
worksheet1.write("C9", "North")
worksheet1.write("C10", "North Total", bold)

worksheet1.write("C5", "Sales", bold)
worksheet1.write("C6", 1000)
worksheet1.write("C7", 1200)
worksheet1.write("C8", 900)
worksheet1.write("C9", 1200)
worksheet1.write("C10", "=SUBTOTAL(9,B2:B5)", bold)



###############################################################################
#
# Example 2: A worksheet with outlined rows. This is the same as the
# previous example except that the rows are collapsed.
# Note: We need to indicate the rows that contains the collapsed symbol '+'
# with the optional parameter, 'collapsed'. The group will be then be
# collapsed if 'hidden' is True.
#
worksheet2.set_row(1, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(2, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(3, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(4, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(5, None, None, {"level": 1, "hidden": True})

worksheet2.set_row(6, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(7, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(8, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(9, None, None, {"level": 2, "hidden": True})
worksheet2.set_row(10, None, None, {"level": 1, "hidden": True})
worksheet2.set_row(11, None, None, {"collapsed": True})

# Adjust the column width for clarity
worksheet2.set_column("A:A", 20)

# Add the data, labels and formulas
worksheet2.write("A1", "Region", bold)
worksheet2.write("A2", "North")
worksheet2.write("A3", "North")
worksheet2.write("A4", "North")
worksheet2.write("A5", "North")
worksheet2.write("A6", "North Total", bold)

worksheet2.write("B1", "Sales", bold)
worksheet2.write("B2", 1000)
worksheet2.write("B3", 1200)
worksheet2.write("B4", 900)
worksheet2.write("B5", 1200)
worksheet2.write("B6", "=SUBTOTAL(9,B2:B5)", bold)

worksheet2.write("A7", "South")
worksheet2.write("A8", "South")
worksheet2.write("A9", "South")
worksheet2.write("A10", "South")
worksheet2.write("A11", "South Total", bold)

worksheet2.write("B7", 400)
worksheet2.write("B8", 600)
worksheet2.write("B9", 500)
worksheet2.write("B10", 600)
worksheet2.write("B11", "=SUBTOTAL(9,B7:B10)", bold)

worksheet2.write("A12", "Grand Total", bold)
worksheet2.write("B12", "=SUBTOTAL(9,B2:B10)", bold)


###############################################################################
#
# Example 3: Create a worksheet with outlined columns.
#
data = [
    ["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"],
    ["North", 50, 20, 15, 25, 65, 80, "=SUM(B2:G2)"],
    ["South", 10, 20, 30, 50, 50, 50, "=SUM(B3:G3)"],
    ["East", 45, 75, 50, 15, 75, 100, "=SUM(B4:G4)"],
    ["West", 15, 15, 55, 35, 20, 50, "=SUM(B5:G5)"],
]

# Add bold format to the first row.
worksheet3.set_row(0, None, bold)

# Set column formatting and the outline level.
worksheet3.set_column("A:A", 10, bold)
worksheet3.set_column("B:G", 5, None, {"level": 1})
worksheet3.set_column("H:H", 10)

# Write the data and a formula
for row, data_row in enumerate(data):
    worksheet3.write_row(row, 0, data_row)

worksheet3.write("H6", "=SUM(H2:H5)", bold)


###############################################################################
#
# Example 4: Show all possible outline levels.
#
levels = [
    "Level 1",
    "Level 2",
    "Level 3",
    "Level 4",
    "Level 5",
    "Level 6",
    "Level 7",
    "Level 6",
    "Level 5",
    "Level 4",
    "Level 3",
    "Level 2",
    "Level 1",
]

worksheet4.write_column("A1", levels)

worksheet4.set_row(0, None, None, {"level": 1})
worksheet4.set_row(1, None, None, {"level": 2})
worksheet4.set_row(2, None, None, {"level": 3})
worksheet4.set_row(3, None, None, {"level": 4})
worksheet4.set_row(4, None, None, {"level": 5})
worksheet4.set_row(5, None, None, {"level": 6})
worksheet4.set_row(6, None, None, {"level": 7})
worksheet4.set_row(7, None, None, {"level": 6})
worksheet4.set_row(8, None, None, {"level": 5})
worksheet4.set_row(9, None, None, {"level": 4})
worksheet4.set_row(10, None, None, {"level": 3})
worksheet4.set_row(11, None, None, {"level": 2})
worksheet4.set_row(12, None, None, {"level": 1})

workbook.close()