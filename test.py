import pandas as pd
from classes import control
from openpyxl import Workbook, utils
from openpyxl.styles import PatternFill, Font, NamedStyle, Alignment
from datetime import datetime
import warnings
import numpy as np

warnings.filterwarnings("ignore")

path_output = './output/'
path_write = './output/tests/'
path_input = './input/'

df_final = pd.read_excel(path_output + 'df_final.xlsx')
# reference_date = control.reference_date
reference_date = datetime(2023,10,1)

df_final['Date'] = reference_date + pd.to_timedelta(df_final['TimeStep'], unit = 'h')
df_final.set_index('Date', inplace = True)

df_final['Year'] = df_final.index.year.astype(int)
df_final['Week'] = df_final.index.week.astype(int)
df_final['Year Month'] = df_final.index.strftime('%Y-%m')
df_final['Year Week'] = df_final.index.strftime('%Y-%W').astype(str)
df_final['Year Week'] = [ s.replace('-', ' - wk ') for s in df_final['Year Week']]

df_final.to_excel(path_write + 'df_final.xlsx')

weekly_sum_df = df_final.groupby('Year Week').sum()
weekly_sum_df.to_excel(path_write + 'weekly_sum_df.xlsx')

monthly_sum_df = df_final.groupby('Year Month').sum()
monthly_sum_df.to_excel(path_write + 'monthly_sum_df.xlsx')

annual_sum_df = df_final.groupby('Year').sum()
annual_sum_df.to_excel(path_write + 'annual_sum_df.xlsx')

# Create a new Excel workbook
wb = Workbook()

# Select the active sheet (you can also create a new one if needed)
sheet = wb.active

#some formats
thousands_separator_format = NamedStyle(name="thousands_separator_format")
thousands_separator_format.number_format = '_-* #,##0_-;-* #,##0_-;_-* "-"??_-;_-@_-'

centered_style = NamedStyle(name="centered_style")
centered_style.alignment = Alignment(horizontal = "center", vertical = "center")

normal_style = NamedStyle(name="normal_style")
normal_style.alignment = Alignment(horizontal = "left", vertical = "center", indent = 0)

normal_style_with_indent = NamedStyle(name="normal_style_with_indent")
normal_style_with_indent.alignment = Alignment(horizontal = "left", vertical = "center", indent = 2)


# function for formatting cells
def formatting_cells(minRow, maxRow, minCol, maxCol, colorCode, boldOption, fontOption, sizeOption, styleOption):
    for row in sheet.iter_rows(min_row = minRow, max_row = maxRow, min_col = minCol, max_col = maxCol):
        for cell in row:
            cell.style = styleOption
            cell.fill = PatternFill(start_color = colorCode, end_color=colorCode, fill_type="solid")
            cell.font = Font(bold = boldOption, size = sizeOption, name = fontOption)
    return


def conditional_formatting_cells(minRow, maxRow, minCol, maxCol, colorCode, boldOption, fontOption, sizeOption, styleOption,fontColor):
    for cell in sheet.iter_rows(min_row = minRow, max_row = maxRow, min_col = minCol, max_col = maxCol):
        for c in cell:
            c.style = styleOption
            c.fill = PatternFill(start_color = colorCode, end_color = colorCode, fill_type="solid")
            if c.value >= 0:
                c.font = Font(bold = boldOption, size = sizeOption, name = fontOption)
            else:
                c.font = Font(bold = boldOption, size = sizeOption, name = fontOption, color = fontColor)
    return

def set_column_width(width, index):
    column_index = index
    column_letter = utils.get_column_letter(column_index)
    column_dimension = sheet.column_dimensions[column_letter]
    column_dimension.width = width
    return

# function for inserting values

# Inserting values and formating cells from the top of the sheet to down
# --------------------------------------------------------------------------
# region top strips

sheet.cell(row = 2, column = 2, value = 'Financial Analysis - Calculations (EoP)')
sheet.cell(row = 3, column = 2, value = 'Units: €')

formatting_cells(minRow = 2, maxRow = 2, minCol = 1, maxCol = 1000, colorCode = "009EE3", 
                 boldOption = True, sizeOption = 11, fontOption = 'Arial', styleOption= normal_style)

formatting_cells(minRow = 3, maxRow = 3, minCol = 1, maxCol = 1000, colorCode = "93D3F7", 
                 boldOption = False, sizeOption = 11, fontOption = 'Arial', styleOption= normal_style)

# endregion
# --------------------------------------------------------------------------
# region inserting operation costs

col_index_yearly_data = 5
col_index_monthly_data = 5 + len(annual_sum_df) + 1
col_index_weekly_data = 5 + len(annual_sum_df) + 1 + len(monthly_sum_df) + 1

# operation costs title
sheet.cell(row = 7, column = 3, value = '1 - Operation Costs')

#formating row operation costs
formatting_cells(minRow = 7, maxRow = 7, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

# Year, month and weeks title
formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_yearly_data, maxCol = col_index_yearly_data + len(annual_sum_df) - 1, colorCode = "D9D9D9", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)
formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_monthly_data, maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, colorCode = "D9D9D9", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)
formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_weekly_data, maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, colorCode = "D9D9D9", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)


# inserting and formatting ANNUAL DATA
list_columns = [s for s in annual_sum_df if 'op_cost' in s]
annual_costs_df = annual_sum_df[list_columns]
annual_costs_df *= -1 

#inserting title of row columns
for idx, value in enumerate(annual_costs_df.columns, start = 8):
    sheet.cell(row = idx, column = 3, value = value)

#inserting years titles
for idx, value in enumerate(annual_costs_df.index, start = 5):
    sheet.cell(row = 5, column = idx, value = value)

#inserting values of yearly data
for row_idx, row in enumerate(annual_costs_df.iterrows(), start = 5):
    for c_idx, value in enumerate(row[1], start = 8):
        sheet.cell(row = c_idx, column = row_idx, value = value)

#formatting titles of row columns
formatting_cells(minRow = 8, maxRow =  8 + len(annual_costs_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption= normal_style_with_indent)

#formatting numbers of yearly values
conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_costs_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                 colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                 fontColor = "FF0000")

# formatting annual sum for the years
list_annual_sum = annual_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_annual_sum, start = 5):
    sheet.cell(row = 7, column = col_idx, value = value)

# formatting annual sum for the years
conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = 5, maxCol = 5 + len(list_annual_sum) - 1, 
                                 colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                 fontColor = "FF0000")


# inserting and formatting MONTHLY DATA
list_columns = [s for s in monthly_sum_df if 'op_cost' in s]
monthly_costs_df = monthly_sum_df[list_columns]
monthly_costs_df *= -1 

#inserting month titles
for idx, value in enumerate(monthly_costs_df.index, start = col_index_monthly_data):
    sheet.cell(row = 5, column = idx, value = value)

for row_idx, row in enumerate(monthly_costs_df.iterrows(), start = col_index_monthly_data):
    for c_idx, value in enumerate(row[1], start = 8):
        sheet.cell(row = c_idx, column = row_idx, value = value)

conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_costs_df.columns) - 1, minCol = col_index_monthly_data, 
                             maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                             colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")

# formatting annual sum for the months
list_monthly_sum = monthly_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_monthly_sum, start = col_index_monthly_data):
    sheet.cell(row = 7, column = col_idx, value = value)

# formatting monthly sum for the months
conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = col_index_monthly_data, 
                             maxCol = col_index_monthly_data + len(list_monthly_sum) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")


# inserting and formatting WEEKLY DATA
list_columns = [s for s in weekly_sum_df if 'op_cost' in s]
weekly_costs_df = weekly_sum_df[list_columns]
weekly_costs_df *= -1 

#inserting week titles
for idx, value in enumerate(weekly_costs_df.index, start = col_index_weekly_data):
    sheet.cell(row = 5, column = idx, value = value)

for row_idx, row in enumerate(weekly_costs_df.iterrows(), start = col_index_weekly_data):
    for c_idx, value in enumerate(row[1], start = 8):
        sheet.cell(row = c_idx, column = row_idx, value = value)

conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_costs_df.columns) - 1, minCol = col_index_weekly_data, 
                             maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                             colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")

# formatting annual sum for the weeks
list_weekly_sum = weekly_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_weekly_sum, start = col_index_weekly_data):
    sheet.cell(row = 7, column = col_idx, value = value)

# formatting monthly sum for the weeks
conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = col_index_weekly_data, 
                             maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")


# endregion
# --------------------------------------------------------------------------
# region investment costs

row_index_inv_data = 7 + 2 + len(annual_costs_df.columns)

# operation costs title
sheet.cell(row = row_index_inv_data, column = 3, value = '2 - Investment Costs')

#formating row operation costs
formatting_cells(minRow = row_index_inv_data, maxRow = row_index_inv_data, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

# inserting and formatting ANNUAL DATA
list_columns = [s for s in annual_sum_df if 'inv_cost' in s]
annual_costs_df = annual_sum_df[list_columns]
annual_costs_df *= -1 

#inserting title of row columns
for idx, value in enumerate(annual_costs_df.columns, start = row_index_inv_data + 1):
    sheet.cell(row = idx, column = 3, value = value)

#inserting values of yearly data
for row_idx, row in enumerate(annual_costs_df.iterrows(), start = 5):
    for c_idx, value in enumerate(row[1], start = row_index_inv_data + 1):
        sheet.cell(row = c_idx, column = row_idx, value = value)

#formatting titles of row columns
formatting_cells(minRow = row_index_inv_data + 1, maxRow =  row_index_inv_data + 1 + len(annual_costs_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style_with_indent)

#formatting numbers of yearly values
conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  = row_index_inv_data + 1  + len(annual_costs_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                 colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                 fontColor = "FF0000")

# formatting annual sum for the years
list_annual_sum_inv = annual_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_annual_sum_inv, start = 5):
    sheet.cell(row = row_index_inv_data, column = col_idx, value = value)

# formatting annual sum for the years
conditional_formatting_cells(minRow = row_index_inv_data, maxRow  = row_index_inv_data, minCol = 5, maxCol = 5 + len(list_annual_sum_inv) - 1, 
                                 colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                 fontColor = "FF0000")



# inserting and formatting MONTHLY DATA
list_columns = [s for s in monthly_sum_df if 'inv_cost' in s]
monthly_costs_df = monthly_sum_df[list_columns]
monthly_costs_df *= -1 

#inserting data for monthly values
for row_idx, row in enumerate(monthly_costs_df.iterrows(), start = col_index_monthly_data):
    for c_idx, value in enumerate(row[1], start = row_index_inv_data + 1 ):
        sheet.cell(row = c_idx, column = row_idx, value = value)

conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  =  row_index_inv_data + 1 + len(annual_costs_df.columns) - 1, minCol = col_index_monthly_data, 
                             maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                             colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")

# inserting annual sum for the months
list_monthly_sum_inv = monthly_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_monthly_sum_inv, start = col_index_monthly_data):
    sheet.cell(row =  row_index_inv_data, column = col_idx, value = value)

# formatting monthly sum for the years
conditional_formatting_cells(minRow =  row_index_inv_data, maxRow  =  row_index_inv_data, minCol = col_index_monthly_data, 
                             maxCol = col_index_monthly_data + len(list_monthly_sum_inv) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")


# inserting and formatting WEEKLY DATA
list_columns = [s for s in weekly_sum_df if 'inv_cost' in s]
weekly_costs_df = weekly_sum_df[list_columns]
weekly_costs_df *= -1 


for row_idx, row in enumerate(weekly_costs_df.iterrows(), start = col_index_weekly_data):
    for c_idx, value in enumerate(row[1], start = row_index_inv_data + 1):
        sheet.cell(row = c_idx, column = row_idx, value = value)

conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  = row_index_inv_data + 1 + len(annual_costs_df.columns) - 1, minCol = col_index_weekly_data, 
                             maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                             colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")

# inserting total sum for the weeks
list_weekly_sum_inv = weekly_costs_df.sum(axis =1)
for col_idx, value in enumerate(list_weekly_sum_inv, start = col_index_weekly_data):
    sheet.cell(row = row_index_inv_data, column = col_idx, value = value)

# formatting total sum for the weeks
conditional_formatting_cells(minRow = row_index_inv_data, maxRow  = row_index_inv_data, minCol = col_index_weekly_data, 
                             maxCol = col_index_weekly_data + len(list_weekly_sum) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")


# endregion
# --------------------------------------------------------------------------
# region other adjustments

set_column_width(20, 3)
set_column_width(3, 4)
set_column_width(3, col_index_monthly_data - 1)
set_column_width(3, col_index_weekly_data -1)

for i in range(col_index_weekly_data, col_index_weekly_data + len(weekly_costs_df)):
    set_column_width(13, i)

# endregion
# --------------------------------------------------------------------------
# region other adjustments

row_index_total = row_index_inv_data + len(annual_costs_df.columns) + 1 + 1

# operation costs title
sheet.cell(row = row_index_total, column = 3, value = '3 - Result')

#formating row operation costs
formatting_cells(minRow = row_index_total, maxRow = row_index_total, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                 boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

annual_result  = np.array(list_annual_sum) + np.array(list_annual_sum_inv)
monthly_result  = np.array(list_monthly_sum) + np.array(list_monthly_sum_inv)
weekly_result  = np.array(list_weekly_sum) + np.array(list_weekly_sum_inv)



# inserting annual result for the years
for col_idx, value in enumerate(annual_result, start = 5):
    sheet.cell(row = row_index_total, column = col_idx, value = value)

# formatting annual sum for the years
conditional_formatting_cells(minRow = row_index_total, maxRow  = row_index_total, minCol = 5, maxCol = 5 + len(list_annual_sum_inv) - 1, 
                                 colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                 fontColor = "FF0000")



# inserting annual sum for the months
for col_idx, value in enumerate(monthly_result, start = col_index_monthly_data):
    sheet.cell(row =  row_index_total, column = col_idx, value = value)

# formatting monthly sum for the years
conditional_formatting_cells(minRow =  row_index_total, maxRow  =  row_index_total, minCol = col_index_monthly_data, 
                             maxCol = col_index_monthly_data + len(list_monthly_sum_inv) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")



# inserting total sum for the weeks
list_weekly_sum_inv = weekly_costs_df.sum(axis =1)
for col_idx, value in enumerate(weekly_result, start = col_index_weekly_data):
    sheet.cell(row = row_index_total, column = col_idx, value = value)

# formatting total sum for the weeks
conditional_formatting_cells(minRow = row_index_total, maxRow  = row_index_total, minCol = col_index_weekly_data, 
                             maxCol = col_index_weekly_data + len(list_weekly_sum) - 1, 
                             colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                             fontColor = "FF0000")

#endregion

# Hide gridlines
sheet.sheet_view.showGridLines = False

# Save the Excel file
wb.save(path_write + "output.xlsx")

