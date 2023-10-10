import pandas as pd
import classes
import sys
import pandas as pd
from classes import control
from openpyxl import Workbook, utils
from openpyxl.styles import PatternFill, Font, NamedStyle, Alignment
import warnings
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime

# def financial_analysis(control):
def financial_analysis():

    warnings.filterwarnings("ignore")

    # path_output = control.path_output
    path_output = './output/'

    df_final = pd.read_excel(path_output + 'df_final.xlsx')
    # reference_date = control.reference_date
    reference_date = datetime(2023,10,10)

    df_final['Date'] = reference_date + pd.to_timedelta(df_final['TimeStep'], unit = 'h')
    df_final.set_index('Date', inplace = True)

    df_final['Year'] = df_final.index.year.astype(int)
    df_final['Week'] = df_final.index.week.astype(int)
    df_final['Year Month'] = df_final.index.strftime('%Y-%m')
    df_final['Year Week'] = df_final.index.strftime('%Y-%W').astype(str)
    df_final['Year Week'] = [ s.replace('-', ' - wk ') for s in df_final['Year Week']]

    weekly_sum_df = df_final.groupby('Year Week').sum()
    monthly_sum_df = df_final.groupby('Year Month').sum()
    annual_sum_df = df_final.groupby('Year').sum()

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

    #defining indexes for columns
    col_index_yearly_data = 5
    col_index_monthly_data = 5 + len(annual_sum_df) + 1
    col_index_weekly_data = 5 + len(annual_sum_df) + 1 + len(monthly_sum_df) + 1

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
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # region inserting REVENUE VALUES

    # operation costs title
    sheet.cell(row = 7, column = 3, value = '1 - Revenue')

    #formating row operation costs
    formatting_cells(minRow = 7, maxRow = 7, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)
    
    # formatting year, month and weeks 
    formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_yearly_data, maxCol = col_index_yearly_data + len(annual_sum_df) - 1, colorCode = "D9D9D9", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)
    formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_monthly_data, maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, colorCode = "D9D9D9", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)
    formatting_cells(minRow = 5, maxRow = 5, minCol = col_index_weekly_data, maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, colorCode = "D9D9D9", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = centered_style)
    
    # inserting and formatting ANNUAL DATA
    list_columns = [s for s in annual_sum_df if '_rev' in s or 'rev_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    annual_revenue_df = annual_sum_df[list_columns]

    #inserting title of row titles
    for idx, value in enumerate(annual_revenue_df.columns, start = 8):
        sheet.cell(row = idx, column = 3, value = value)

    #inserting years titles
    for idx, value in enumerate(annual_revenue_df.index, start = 5):
        sheet.cell(row = 5, column = idx, value = value)

    #inserting values of yearly data
    for row_idx, row in enumerate(annual_revenue_df.iterrows(), start = 5):
        for c_idx, value in enumerate(row[1], start = 8):
            sheet.cell(row = c_idx, column = row_idx, value = value)


    #formatting titles of row columns
    formatting_cells(minRow = 8, maxRow =  8 + len(annual_revenue_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption= normal_style_with_indent)

    #formatting numbers of yearly values
    conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_revenue_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                    colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")

    # formatting annual sum for the years
    list_annual_sum_revenue = annual_revenue_df.sum(axis =1)
    for col_idx, value in enumerate(list_annual_sum_revenue, start = 5):
        sheet.cell(row = 7, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = 5, maxCol = 5 + len(list_annual_sum_revenue) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")
    





    # inserting and formatting MONTHLY DATA
    list_columns = [s for s in monthly_sum_df if '_rev' in s or 'rev_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    monthly_revenue_df = monthly_sum_df[list_columns]

    #inserting month titles
    for idx, value in enumerate(monthly_revenue_df.index, start = col_index_monthly_data):
        sheet.cell(row = 5, column = idx, value = value)

    for row_idx, row in enumerate(monthly_revenue_df.iterrows(), start = col_index_monthly_data):
        for c_idx, value in enumerate(row[1], start = 8):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_revenue_df.columns) - 1, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")
    

   # formatting annual sum for the months
    list_monthly_sum_revenue = monthly_revenue_df.sum(axis =1)
    for col_idx, value in enumerate(list_monthly_sum_revenue, start = col_index_monthly_data):
        sheet.cell(row = 7, column = col_idx, value = value)

    # formatting monthly sum for the months
    conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_revenue) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")





    # inserting and formatting WEEKLY DATA
    list_columns = [s for s in weekly_sum_df if '_rev' in s or 'rev_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    weekly_revenue_df = weekly_sum_df[list_columns]

    #inserting week titles
    for idx, value in enumerate(weekly_revenue_df.index, start = col_index_weekly_data):
        sheet.cell(row = 5, column = idx, value = value)

    for row_idx, row in enumerate(weekly_revenue_df.iterrows(), start = col_index_weekly_data):
        for c_idx, value in enumerate(row[1], start = 8):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = 8, maxRow  = 8 + len(annual_revenue_df.columns) - 1, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")
    
    # formatting annual sum for the weeks
    list_weekly_sum_revenue = weekly_revenue_df.sum(axis =1)
    for col_idx, value in enumerate(list_weekly_sum_revenue, start = col_index_weekly_data):
        sheet.cell(row = 7, column = col_idx, value = value)

    # formatting monthly sum for the weeks
    conditional_formatting_cells(minRow = 7, maxRow  = 7, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")


    # endregion
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # region inserting REVENUE VALUES COMPENSATION


    row_index_comp_data = 7 + 2 + len(annual_revenue_df.columns)

    # operation costs title
    sheet.cell(row = row_index_comp_data, column = 3, value = '2 - Revenue (compensation)')

    #formating row operation costs
    formatting_cells(minRow = row_index_comp_data, maxRow = row_index_comp_data, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

    # inserting and formatting ANNUAL DATA
    list_columns = [s for s in annual_sum_df if 'comp_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    annual_comp_df = annual_sum_df[list_columns]

    #inserting title of row columns
    for idx, value in enumerate(annual_comp_df.columns, start = row_index_comp_data + 1):
        sheet.cell(row = idx, column = 3, value = value)

    #inserting values of yearly data
    for row_idx, row in enumerate(annual_comp_df.iterrows(), start = 5):
        for c_idx, value in enumerate(row[1], start = row_index_comp_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    #formatting titles of row columns
    formatting_cells(minRow = row_index_comp_data + 1, maxRow =  row_index_comp_data + 1 + len(annual_comp_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style_with_indent)

    #formatting numbers of yearly values
    conditional_formatting_cells(minRow = row_index_comp_data + 1 , maxRow  = row_index_comp_data + 1  + len(annual_comp_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                    colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")

    # formatting annual sum for the years
    list_annual_sum_compensation = annual_comp_df.sum(axis =1)
    for col_idx, value in enumerate(list_annual_sum_compensation, start = 5):
        sheet.cell(row = row_index_comp_data, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = row_index_comp_data, maxRow  = row_index_comp_data, minCol = 5, maxCol = 5 + len(list_annual_sum_compensation) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")





    # inserting and formatting MONTHLY DATA
    list_columns = [s for s in monthly_sum_df if 'comp_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    monthly_comp_df = monthly_sum_df[list_columns]

    #inserting data for monthly values
    for row_idx, row in enumerate(monthly_comp_df.iterrows(), start = col_index_monthly_data):
        for c_idx, value in enumerate(row[1], start = row_index_comp_data + 1 ):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_comp_data + 1 , maxRow  =  row_index_comp_data + 1 + len(annual_comp_df.columns) - 1, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting annual sum for the months
    list_monthly_sum_compensation = monthly_comp_df.sum(axis =1)
    for col_idx, value in enumerate(list_monthly_sum_compensation, start = col_index_monthly_data):
        sheet.cell(row =  row_index_comp_data, column = col_idx, value = value)

    # formatting monthly sum for the years
    conditional_formatting_cells(minRow =  row_index_comp_data, maxRow  =  row_index_comp_data, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_compensation) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")




    # inserting and formatting WEEKLY DATA
    list_columns = [s for s in weekly_sum_df if 'comp_' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    weekly_comp_df = weekly_sum_df[list_columns]


    for row_idx, row in enumerate(weekly_comp_df.iterrows(), start = col_index_weekly_data):
        for c_idx, value in enumerate(row[1], start = row_index_comp_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_comp_data + 1 , maxRow  = row_index_comp_data + 1 + len(annual_comp_df.columns) - 1, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting total sum for the weeks
    list_weekly_sum_compensation = weekly_comp_df.sum(axis =1)
    for col_idx, value in enumerate(list_weekly_sum_compensation, start = col_index_weekly_data):
        sheet.cell(row = row_index_comp_data, column = col_idx, value = value)

    # formatting total sum for the weeks
    conditional_formatting_cells(minRow = row_index_comp_data, maxRow  = row_index_comp_data, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(list_weekly_sum_compensation) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")
    

    # endregion
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # region inserting OPERATION COSTS

    row_index_op_costs_data = 7 + 4 + len(annual_revenue_df.columns) + len(annual_comp_df.columns)

    # operation costs title
    sheet.cell(row = row_index_op_costs_data, column = 3, value = '3 - Operational costs')

    #formating row operation costs
    formatting_cells(minRow = row_index_op_costs_data, maxRow = row_index_op_costs_data, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

    # inserting and formatting ANNUAL DATA
    list_columns = [s for s in annual_sum_df if 'op_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    annual_op_costs_df = annual_sum_df[list_columns]
    annual_op_costs_df *= -1

    #inserting title of row columns
    for idx, value in enumerate(annual_op_costs_df.columns, start = row_index_op_costs_data + 1):
        sheet.cell(row = idx, column = 3, value = value)

    #inserting values of yearly data
    for row_idx, row in enumerate(annual_op_costs_df.iterrows(), start = 5):
        for c_idx, value in enumerate(row[1], start = row_index_op_costs_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    #formatting titles of row columns
    formatting_cells(minRow = row_index_op_costs_data + 1, maxRow =  row_index_op_costs_data + 1 + len(annual_op_costs_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style_with_indent)

    #formatting numbers of yearly values
    conditional_formatting_cells(minRow = row_index_op_costs_data + 1 , maxRow  = row_index_op_costs_data + 1  + len(annual_op_costs_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                    colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")

    # formatting annual sum for the years
    list_annual_sum_op_costs = annual_op_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_annual_sum_op_costs, start = 5):
        sheet.cell(row = row_index_op_costs_data, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = row_index_op_costs_data, maxRow  = row_index_op_costs_data, minCol = 5, maxCol = 5 + len(list_annual_sum_op_costs) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")




    # inserting and formatting MONTHLY DATA
    list_columns = [s for s in monthly_sum_df if 'op_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    monthly_op_costs_df = monthly_sum_df[list_columns]
    monthly_op_costs_df *= -1

    #inserting data for monthly values
    for row_idx, row in enumerate(monthly_op_costs_df.iterrows(), start = col_index_monthly_data):
        for c_idx, value in enumerate(row[1], start = row_index_op_costs_data+ 1 ):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_op_costs_data + 1 , maxRow  =  row_index_op_costs_data + 1 + len(annual_op_costs_df.columns) - 1, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting annual sum for the months
    list_monthly_sum_op_costs = monthly_op_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_monthly_sum_op_costs, start = col_index_monthly_data):
        sheet.cell(row =  row_index_op_costs_data, column = col_idx, value = value)

    # formatting monthly sum for the years
    conditional_formatting_cells(minRow =  row_index_op_costs_data, maxRow  =  row_index_op_costs_data, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_op_costs) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")




    # inserting and formatting WEEKLY DATA
    list_columns = [s for s in weekly_sum_df if 'op_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    weekly_op_costs_df = weekly_sum_df[list_columns]
    weekly_op_costs_df *= -1


    for row_idx, row in enumerate(weekly_op_costs_df.iterrows(), start = col_index_weekly_data):
        for c_idx, value in enumerate(row[1], start = row_index_op_costs_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_op_costs_data + 1 , maxRow  = row_index_op_costs_data + 1 + len(annual_op_costs_df.columns) - 1, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting total sum for the weeks
    list_weekly_sum_op_costs = weekly_op_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_weekly_sum_op_costs, start = col_index_weekly_data):
        sheet.cell(row = row_index_op_costs_data, column = col_idx, value = value)

    # formatting total sum for the weeks
    conditional_formatting_cells(minRow = row_index_op_costs_data, maxRow  = row_index_op_costs_data, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(list_weekly_sum_op_costs) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")


    # endregion
    # ---------------------------------------------------------------------------------------------------------------------------------------
    # region inserting investment COSTS

    row_index_inv_data = 7 + 6 + len(annual_revenue_df.columns) + len(annual_comp_df.columns) + len(annual_op_costs_df.columns)

    # operation costs title
    sheet.cell(row = row_index_inv_data, column = 3, value = '4 - Investment costs')

    #formating row operation costs
    formatting_cells(minRow = row_index_inv_data, maxRow = row_index_inv_data, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

    # inserting and formatting ANNUAL DATA
    list_columns = [s for s in annual_sum_df if 'inv_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    annual_inv_costs_df = annual_sum_df[list_columns]
    annual_inv_costs_df *= -1

    #inserting title of row columns
    for idx, value in enumerate(annual_inv_costs_df.columns, start = row_index_inv_data + 1):
        sheet.cell(row = idx, column = 3, value = value)

    #inserting values of yearly data
    for row_idx, row in enumerate(annual_inv_costs_df.iterrows(), start = 5):
        for c_idx, value in enumerate(row[1], start = row_index_inv_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    #formatting titles of row columns
    formatting_cells(minRow = row_index_inv_data + 1, maxRow =  row_index_inv_data + 1 + len(annual_inv_costs_df.columns) - 1, minCol = 1, maxCol = 3, colorCode = "D4EEFC", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style_with_indent)

    #formatting numbers of yearly values
    conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  = row_index_inv_data + 1  + len(annual_inv_costs_df.columns) - 1, minCol = 5, maxCol = 5 + len(annual_sum_df) - 1, 
                                    colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")

    # formatting annual sum for the years
    list_annual_sum_inv_costs = annual_inv_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_annual_sum_inv_costs, start = 5):
        sheet.cell(row = row_index_inv_data, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = row_index_inv_data, maxRow  = row_index_inv_data, minCol = 5, maxCol = 5 + len(list_annual_sum_inv_costs) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")





    # inserting and formatting MONTHLY DATA
    list_columns = [s for s in monthly_sum_df if 'inv_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    monthly_inv_costs_df = monthly_sum_df[list_columns]
    monthly_inv_costs_df *= -1

    #inserting data for monthly values
    for row_idx, row in enumerate(monthly_inv_costs_df.iterrows(), start = col_index_monthly_data):
        for c_idx, value in enumerate(row[1], start = row_index_inv_data+ 1 ):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  =  row_index_inv_data + 1 + len(annual_inv_costs_df.columns) - 1, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(monthly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting annual sum for the months
    list_monthly_sum_inv_costs = monthly_inv_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_monthly_sum_inv_costs, start = col_index_monthly_data):
        sheet.cell(row =  row_index_inv_data, column = col_idx, value = value)

    # formatting monthly sum for the years
    conditional_formatting_cells(minRow =  row_index_inv_data, maxRow  =  row_index_inv_data, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_inv_costs) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")





    # inserting and formatting WEEKLY DATA
    list_columns = [s for s in weekly_sum_df if 'inv_costs' in s]
    list_columns = [s for s in list_columns if 'total' not in s]
    list_columns.sort()
    weekly_inv_costs_df = weekly_sum_df[list_columns]
    weekly_inv_costs_df *= -1


    for row_idx, row in enumerate(weekly_inv_costs_df.iterrows(), start = col_index_weekly_data):
        for c_idx, value in enumerate(row[1], start = row_index_inv_data + 1):
            sheet.cell(row = c_idx, column = row_idx, value = value)

    conditional_formatting_cells(minRow = row_index_inv_data + 1 , maxRow  = row_index_inv_data + 1 + len(annual_inv_costs_df.columns) - 1, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(weekly_sum_df) - 1, 
                                colorCode = "D4EEFC", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting total sum for the weeks
    list_weekly_sum_inv_costs = weekly_inv_costs_df.sum(axis =1)
    for col_idx, value in enumerate(list_weekly_sum_inv_costs, start = col_index_weekly_data):
        sheet.cell(row = row_index_inv_data, column = col_idx, value = value)

    # formatting total sum for the weeks
    conditional_formatting_cells(minRow = row_index_inv_data, maxRow  = row_index_inv_data, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(list_weekly_sum_inv_costs) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # endregion
    # --------------------------------------------------------------------------
    # region other adjustments


    set_column_width(23, 3)
    set_column_width(2, 4)
    set_column_width(2, col_index_monthly_data - 1)
    set_column_width(2, col_index_weekly_data -1)

    for i in range(col_index_weekly_data, col_index_weekly_data + len(weekly_comp_df)):
        set_column_width(13, i)


    # endregion
    # --------------------------------------------------------------------------
    # region FINAL SUM RESULT


    row_index_total = 7 + 8 + len(annual_revenue_df.columns) + len(annual_comp_df.columns) + len(annual_op_costs_df.columns) +len(annual_inv_costs_df.columns)

    # operation costs title
    sheet.cell(row = row_index_total, column = 3, value = '5 - Result')

    #formating row operation costs
    formatting_cells(minRow = row_index_total, maxRow = row_index_total, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)

    annual_result  = np.array(list_annual_sum_revenue) + np.array(list_annual_sum_compensation) + np.array(list_annual_sum_op_costs) + np.array(list_annual_sum_inv_costs)
    monthly_result  = np.array(list_monthly_sum_revenue) + np.array(list_monthly_sum_compensation) + np.array(list_monthly_sum_op_costs) + np.array(list_monthly_sum_inv_costs)
    weekly_result  = np.array(list_weekly_sum_revenue) + np.array(list_weekly_sum_compensation) + np.array(list_weekly_sum_op_costs) + np.array(list_weekly_sum_inv_costs)

    # inserting annual result for the years
    for col_idx, value in enumerate(annual_result, start = 5):
        sheet.cell(row = row_index_total, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = row_index_total, maxRow  = row_index_total, minCol = 5, maxCol = 5 + len(list_annual_sum_revenue) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")


    # inserting annual sum for the months
    for col_idx, value in enumerate(monthly_result, start = col_index_monthly_data):
        sheet.cell(row =  row_index_total, column = col_idx, value = value)

    # formatting monthly sum for the years
    conditional_formatting_cells(minRow =  row_index_total, maxRow  =  row_index_total, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_revenue) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")


    # inserting total sum for the weeks
    for col_idx, value in enumerate(weekly_result, start = col_index_weekly_data):
        sheet.cell(row = row_index_total, column = col_idx, value = value)

    # formatting total sum for the weeks
    conditional_formatting_cells(minRow = row_index_total, maxRow  = row_index_total, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(list_weekly_sum_revenue) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")
    

    # now inserting the cumulated sum of the cash flows
    row_index_total_cumulated = row_index_total + 1 + 1

    # operation costs title
    sheet.cell(row = row_index_total_cumulated, column = 3, value = '6 - Result accumulated')

    #formating row operation costs
    formatting_cells(minRow = row_index_total_cumulated, maxRow = row_index_total_cumulated, minCol = 1, maxCol = 3, colorCode = "93D3F7", 
                    boldOption = True, sizeOption = 10, fontOption = 'Arial', styleOption = normal_style)
    
    
    # inserting annual result for the years
    for col_idx, value in enumerate(np.cumsum(annual_result), start = 5):
        sheet.cell(row = row_index_total_cumulated, column = col_idx, value = value)

    # formatting annual sum for the years
    conditional_formatting_cells(minRow = row_index_total_cumulated, maxRow  = row_index_total_cumulated, minCol = 5, maxCol = 5 + len(list_annual_sum_revenue) - 1, 
                                    colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                    fontColor = "FF0000")

    # inserting annual sum for the months
    for col_idx, value in enumerate(np.cumsum(monthly_result), start = col_index_monthly_data):
        sheet.cell(row =  row_index_total_cumulated, column = col_idx, value = value)

    # formatting monthly sum for the years
    conditional_formatting_cells(minRow =  row_index_total_cumulated, maxRow  =  row_index_total_cumulated, minCol = col_index_monthly_data, 
                                maxCol = col_index_monthly_data + len(list_monthly_sum_revenue) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")

    # inserting total sum for the weeks
    for col_idx, value in enumerate(np.cumsum(weekly_result), start = col_index_weekly_data):
        sheet.cell(row = row_index_total_cumulated, column = col_idx, value = value)

    # formatting total sum for the weeks
    conditional_formatting_cells(minRow = row_index_total_cumulated, maxRow  = row_index_total_cumulated, minCol = col_index_weekly_data, 
                                maxCol = col_index_weekly_data + len(list_weekly_sum_revenue) - 1, 
                                colorCode = "93D3F7", boldOption = True, fontOption = 'Arial', sizeOption = 10 , styleOption = thousands_separator_format,
                                fontColor = "FF0000")
    
    #endregion

    # Hide gridlines
    sheet.sheet_view.showGridLines = False
    mycell = sheet['D6']
    sheet.freeze_panes = mycell

    # Save the Excel file
    wb.save(path_output + "financial_analysis.xlsx")


financial_analysis()