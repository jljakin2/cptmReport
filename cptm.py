import random
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import numpy as np
import time
import xlsxwriter


# global variables
hr_filename = ""
p59_filename_us = ""
p59_filename_canada = ""
p59_filename_mexico = ""

# functions
def run():
    # Define variables for all files needed. Inputs come from frontend GUI
    hr_global_file = hr_filename
    p59_file_us = p59_filename_us
    p59_file_canada = p59_filename_canada
    p59_file_mexico = p59_filename_mexico
    # Define HR Global dataframes for US and clean where necessary
    hr_df_us = pd.read_excel(hr_global_file, sheet_name=2)
    hr_df_us_dup = hr_df_us.drop_duplicates(subset='User')
    hr_df_us_clean = hr_df_us_dup.replace([0], np.nan)
    hr_df_us_clean.rename(columns = {'User':'Global Id'}, inplace=True)

    # Define HR Global dataframes for Canada and clean where necessary
    hr_df_canada = pd.read_excel(hr_global_file, sheet_name=0)
    hr_df_canada_dup = hr_df_canada.drop_duplicates(subset='User')
    hr_df_canada_clean = hr_df_canada_dup.replace([0], np.nan)
    hr_df_canada_clean.rename(columns = {'User':'Global Id'}, inplace=True)

    # Define HR Global dataframes for Mexico  and clean where necessary
    hr_df_mexico = pd.read_excel(hr_global_file, sheet_name=1)
    hr_df_mexico_dup = hr_df_mexico.drop_duplicates(subset='User')
    hr_df_mexico_clean = hr_df_mexico_dup.replace([0], np.nan)
    hr_df_mexico_clean.rename(columns = {'User':'Global Id'}, inplace=True)

    # Define P59 dataframes. Raw data is clean.
    p59_df_us = pd.read_excel(p59_file_us)
    p59_df_canada = pd.read_excel(p59_file_canada)
    p59_df_mexico = pd.read_excel(p59_file_mexico)

    # Combine dataframes based on country
    combined_us = pd.merge(hr_df_us_clean, p59_df_us, on='Global Id', how = 'outer')
    combined_canada = pd.merge(hr_df_canada_clean, p59_df_canada, on='Global Id', how = 'outer')
    combined_mexico = pd.merge(hr_df_mexico_clean, p59_df_mexico, on='Global Id', how = 'outer')

    # Percentage for US
    total_us = combined_us["Global Id"].count()
    curr_total_us = combined_us["Completed"].count()
    total_curr_perc_us = curr_total_us / total_us

    # Percentage for US grouped by indirect/direct
    total_DI_us = combined_us.groupby("Dir/Ind Labor")["Global Id"].count()
    curr_DI_us = combined_us.groupby("Direct/Indirect")["Global Id"].count()
    DI_curr_perc_us = curr_DI_us / total_DI_us

    # Percentage for US grouped by division
    total_div_us = combined_us.groupby("GB Division")["Global Id"].count()
    curr_div_us = combined_us.groupby("GB Division")["Completed"].count()
    div_curr_perc_us = curr_div_us / total_div_us

    # Percentage for Canada
    total_canada = combined_canada["Global Id"].count()
    curr_total_canada = combined_canada["Completed"].count()
    total_curr_perc_canada = curr_total_canada / total_canada

    # Percentage for Canada grouped by indirect/direct
    total_DI_canada = combined_canada.groupby("Dir/Ind Labor")["Global Id"].count()
    curr_DI_canada = combined_canada.groupby("Direct/Indirect")["Global Id"].count()
    DI_curr_perc_canada = curr_DI_canada / total_DI_canada

    # Percentage for Canada grouped by division
    total_div_canada = combined_canada.groupby("GB Division")["Global Id"].count()
    curr_div_canada = combined_canada.groupby("GB Division")["Completed"].count()
    div_curr_perc_canada = curr_div_canada / total_div_canada

    # Percentage for Mexico
    total_mexico = combined_mexico["Global Id"].count()
    curr_total_mexico = combined_mexico["Completed"].count()
    total_curr_perc_mexico = curr_total_mexico / total_mexico

    # Percentage for Mexico grouped by indirect/direct
    total_DI_mexico = pd.DataFrame(combined_mexico.groupby("Dir/Ind Labor")["Global Id"].count())
    curr_DI_mexico = pd.DataFrame(combined_mexico.groupby("Direct/Indirect")["Global Id"].count())
    DI_curr_perc_mexico = curr_DI_mexico / total_DI_mexico

    # Percentage for Mexico grouped by division
    total_div_mexico = combined_mexico.groupby("GB Division")["Global Id"].count()
    curr_div_mexico = combined_mexico.groupby("GB Division")["Completed"].count()
    div_curr_perc_mexico = curr_div_mexico / total_div_mexico

    # Total percentages
    total = {'US':[total_curr_perc_us], 'Canada':[total_curr_perc_canada], 'Mexico':[total_curr_perc_mexico]}
    total_df = pd.DataFrame(total, index=['Percentages'])

    # Combined percentages for direct/indirect
    total_DI_curr_perc = pd.concat([DI_curr_perc_us, DI_curr_perc_canada, DI_curr_perc_mexico], axis=1)

    # Total percentages for direct/indirect
    total_DI = pd.concat([total_DI_us, total_DI_canada, total_DI_mexico], axis=1)
    total_DI_total = total_DI.sum(axis=1, skipna=True)
    total_curr_DI = pd.concat([curr_DI_us, curr_DI_canada, curr_DI_mexico], axis=1)
    total_curr_DI_total = total_curr_DI.sum(axis=1, skipna=True)
    total_curr_DI_perc = total_curr_DI_total / total_DI_total

    # Define today's date so it can be added to the name of the output excel file
    todaysdate = time.strftime("%d-%m-%Y")

    # Define filename of output excel file
    excelfilename = "CptM Assigned Curricula Data_" + todaysdate + ".xlsx"
    
    # Define method to export dataframes
    writer = pd.ExcelWriter(excelfilename, engine='xlsxwriter')

    # Write dataframes to corresponding sheets
    total_df.to_excel(writer, sheet_name ='Total') 
    total_DI_curr_perc.to_excel(writer, sheet_name='Direct_Indirect', startrow=0, startcol=0)
    total_curr_DI_perc.to_excel(writer, sheet_name='Direct_Indirect', startrow=0, startcol=7)
    div_curr_perc_us.to_excel(writer, sheet_name='Division', startrow=0, startcol=0)
    div_curr_perc_canada.to_excel(writer, sheet_name='Division', startrow=0, startcol=3)
    div_curr_perc_mexico.to_excel(writer, sheet_name='Division', startrow=0, startcol=6)

    # Create xlsxwriter objects for formatting, worksheet, and chart creation for Total sheet
    total_workbook_object= writer.book
    total_worksheet_object = writer.sheets['Total']
    total_chart_object = total_workbook_object.add_chart({'type': 'column'})
    format_object1 = total_workbook_object.add_format({'num_format': '0 %'})
    total_worksheet_object.set_column('A:D', 11, format_object1)
    # [sheetname, first_row, first_col, last_row, last_col]
    total_chart_object.add_series({ 
        'name':       ['Total', 3, 0],   
        'categories': ['Total', 0, 1, 0, 3],   
        'values':     ['Total', 1, 1, 1, 3],   
        }) 
    total_chart_object.set_title({'name': 'Total Curricula Assigned'}) 
    total_chart_object.set_x_axis({'name': 'Country'})  
    total_chart_object.set_y_axis({'name': 'Percent'})
    total_chart_object.set_legend({'position': 'none'})
    total_chart_object.set_y_axis({'major_gridlines': {'visible': False}})
    total_worksheet_object.insert_chart('F2', total_chart_object, {'x_offset': 20, 'y_offset': 5})

    # Create xlsxwriter objects for formatting, worksheet, and chart creation for Direct/Indirect Sheet
    dir_workbook_object= writer.book
    dir_worksheet_object = writer.sheets['Direct_Indirect']
    dir_chart_object = dir_workbook_object.add_chart({'type': 'column'})
    dir_chart_object_2 = dir_workbook_object.add_chart({'type': 'column'})
    format_object1 = dir_workbook_object.add_format({'num_format': '0 %'})
    dir_worksheet_object.set_column('A:I', 13, format_object1)
    cell_format = dir_workbook_object.add_format({'bold': True, 'align': 'right'})
    dir_worksheet_object.write('B1', 'US', cell_format)
    dir_worksheet_object.write('C1', 'Canada', cell_format)
    dir_worksheet_object.write('D1', 'Mexico', cell_format)
    dir_worksheet_object.write('A2', 'Direct', cell_format)
    dir_worksheet_object.write('A3', 'Indirect', cell_format)
    dir_worksheet_object.write('I1', 'Total', cell_format)
    dir_worksheet_object.write('H2', 'Direct', cell_format)
    dir_worksheet_object.write('H3', 'Indirect', cell_format)
    # [sheetname, first_row, first_col, last_row, last_col]
    dir_chart_object.add_series({ 
        'name':       ['Direct_Indirect', 1, 0],   
        'categories': ['Direct_Indirect', 0, 1, 0, 3],   
        'values':     ['Direct_Indirect', 1, 1, 1, 3],
        })
    dir_chart_object.add_series({ 
        'name':       ['Direct_Indirect', 2, 0],   
        'categories': ['Direct_Indirect', 0, 1, 0, 3],   
        'values':     ['Direct_Indirect', 2, 1, 2, 3],   
        }) 
    dir_chart_object.set_title({'name': 'Curricula Assigned - Direct vs. Indirect'}) 
    dir_chart_object.set_x_axis({'name': 'Country'})  
    dir_chart_object.set_y_axis({'name': 'Percent'})
    dir_chart_object.set_y_axis({'major_gridlines': {'visible': False}})
    dir_worksheet_object.insert_chart('A5', dir_chart_object, {'x_offset': 20, 'y_offset': 5})
    # [sheetname, first_row, first_col, last_row, last_col]
    dir_chart_object_2.add_series({ 
        'name':       ['Direct_Indirect', 0, 8],   
        'categories': ['Direct_Indirect', 1, 7, 2, 7],   
        'values':     ['Direct_Indirect', 1, 8, 2, 8],   
        })
    dir_chart_object_2.set_title({'name': 'Total Curricula Assigned - Direct vs. Indirect'}) 
    dir_chart_object_2.set_x_axis({'name': 'Direct/Indirect'})  
    dir_chart_object_2.set_y_axis({'name': 'Percent'})
    dir_chart_object_2.set_legend({'position': 'none'})
    dir_chart_object_2.set_y_axis({'major_gridlines': {'visible': False}})
    dir_worksheet_object.insert_chart('H5', dir_chart_object_2, {'x_offset': 20, 'y_offset': 5})

    # Create xlsxwriter objects for formatting, worksheet, and chart creation for Division Sheet
    div_workbook_object= writer.book
    div_worksheet_object = writer.sheets['Division']
    div_chart_object_us = div_workbook_object.add_chart({'type': 'column'})
    div_chart_object_canada = div_workbook_object.add_chart({'type': 'column'})
    div_chart_object_mexico = div_workbook_object.add_chart({'type': 'column'})
    format_object3 = dir_workbook_object.add_format({'num_format': '0 %'})
    div_worksheet_object.set_column('A:H', 13, format_object3)
    cell_format = div_workbook_object.add_format({'bold': True, 'align': 'right'})
    div_worksheet_object.write('B1', 'US', cell_format)
    div_worksheet_object.write('E1', 'Canada', cell_format)
    div_worksheet_object.write('H1', 'Mexico', cell_format)
    # [sheetname, first_row, first_col, last_row, last_col]
    div_chart_object_us.add_series({ 
        'name':       ['Division', 0, 1],   
        'categories': ['Division', 1, 0, 33, 0],   
        'values':     ['Division', 1, 1, 33, 1],
        })
    div_chart_object_canada.add_series({ 
        'name':       ['Division', 0, 4],   
        'categories': ['Division', 1, 3, 11, 3],   
        'values':     ['Division', 1, 4, 11, 4],   
        })
    div_chart_object_mexico.add_series({ 
        'name':       ['Division', 0, 7],   
        'categories': ['Division', 1, 6, 25, 6],   
        'values':     ['Division', 1, 7, 25, 7],   
        }) 
    div_chart_object_us.set_title({'name': 'Curricula Assigned - Division'}) 
    div_chart_object_us.set_x_axis({'name': 'United States'})  
    div_chart_object_us.set_y_axis({'name': 'Percent'})
    div_chart_object_us.set_y_axis({'major_gridlines': {'visible': False}})
    div_chart_object_us.set_legend({'position': 'none'})
    div_worksheet_object.insert_chart('K1', div_chart_object_us, {'x_offset': 20, 'y_offset': 5})

    div_chart_object_canada.set_title({'name': 'Curricula Assigned - Division'}) 
    div_chart_object_canada.set_x_axis({'name': 'Canada'})  
    div_chart_object_canada.set_y_axis({'name': 'Percent'})
    div_chart_object_canada.set_y_axis({'major_gridlines': {'visible': False}})
    div_chart_object_canada.set_legend({'position': 'none'})
    div_worksheet_object.insert_chart('K17', div_chart_object_canada, {'x_offset': 20, 'y_offset': 5})

    div_chart_object_mexico.set_title({'name': 'Curricula Assigned - Division'}) 
    div_chart_object_mexico.set_x_axis({'name': 'Mexico'})  
    div_chart_object_mexico.set_y_axis({'name': 'Percent'})
    div_chart_object_mexico.set_y_axis({'major_gridlines': {'visible': False}})
    div_chart_object_mexico.set_legend({'position': 'none'})
    div_worksheet_object.insert_chart('K34', div_chart_object_mexico, {'x_offset': 20, 'y_offset': 5})

    writer.save()



def hr_open():
    global hr_filename
    hr_filename = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    hr_data = Label(root, text=hr_filename, width=40, borderwidth=2, relief="groove")
    hr_data.grid(row=0, column=1)

def p59_open_us():
    global p59_filename_us
    p59_filename_us = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    p59_data_us = Label(root, text=p59_filename_us, width=40, borderwidth=2, relief="groove")
    p59_data_us.grid(row=1, column=1)

def p59_open_canada():
    global p59_filename_canada
    p59_filename_canada = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    p59_data_canada = Label(root, text=p59_filename_canada, width=40, borderwidth=2, relief="groove")
    p59_data_canada.grid(row=2, column=1)

def p59_open_mexico():
    global p59_filename_mexico
    p59_filename_mexico = filedialog.askopenfilename(
        initialdir="/Users/jeffjakinovich/Desktop",
        title="Select a file",
        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")),
    )
    p59_data_mexico = Label(root, text=p59_filename_mexico, width=40, borderwidth=2, relief="groove")
    p59_data_mexico.grid(row=3, column=1)


#list of compliments
happy=[
    "You are the most perfect you there is!",
    "Your perspective is refreshing!",
    "You should be proud of yourself!",
    "In high school, I bet you were voted Most Likely To Keep Being Awesome!",
    "You are really something special!",
    "You look great today!",
    "You are strong!",
    "You light up the room!",
    "Everything you touch turns to gold!",
]


# window setup
root = Tk()
root.geometry("640x235")
root.wm_title("CptM Curricula Data")

# labels
hr_blank = Label(root, width=40, borderwidth=2, relief="groove")
hr_blank.grid(row=0, column=1)

p59_blank_us = Label(root, width=40, borderwidth=2, relief="groove")
p59_blank_us.grid(row=1, column=1)

p59_blank_canada = Label(root, width=40, borderwidth=2, relief="groove")
p59_blank_canada.grid(row=2, column=1)

p59_blank_mexico = Label(root, width=40, borderwidth=2, relief="groove")
p59_blank_mexico.grid(row=3, column=1)

hr_label = Label(root, text="HR Global File")
hr_label.grid(row=0, column=0, padx=10, pady=10)

p59_label_us = Label(root, text="P59 File for US")
p59_label_us.grid(row=1, column=0, padx=10, pady=10)

p59_label_canada = Label(root, text="P59 File for Canada")
p59_label_canada.grid(row=2, column=0, padx=10, pady=10)

p59_label_mexico = Label(root, text="P59 File for Mexico")
p59_label_mexico.grid(row=3, column=0, padx=10, pady=10)

pos_label = Label(root, text=random.choice(happy), wraplength=250, justify=CENTER, borderwidth=2, relief="solid")
pos_label.grid(row=4, column=0, padx=10, pady=10, rowspan=2)

# buttons
run = Button(root, text="Run", width=10, bg="gray60", fg="black", command=run)
run.grid(row=4, column=2, padx=10, pady=10)

import1 = Button(root, text="Import", width=10, command=hr_open)
import1.grid(row=0, column=2, padx=10, pady=10)

import2 = Button(root, text="Import", width=10, command=p59_open_us)
import2.grid(row=1, column=2, padx=10, pady=10)

import3 = Button(root, text="Import", width=10, command=p59_open_canada)
import3.grid(row=2, column=2, padx=10, pady=10)

import4 = Button(root, text="Import", width=10, command=p59_open_mexico)
import4.grid(row=3, column=2, padx=10, pady=10)

root.mainloop()
