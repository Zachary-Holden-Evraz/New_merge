# RPT Merge
# Note that this was created in Python and made into a windows executable file

import os, sys
from sys import exit
import fnmatch
import shutil
import pandas as pd
from threading import Thread
import openpyxl as pyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import *
from tkinter import simpledialog
from tkinter.filedialog import askopenfilename

# Ask the User to give the files - this will allow this to work on any Windows computer
main_file = askopenfilename(title = 'File you want to update') # Excel File we want to add all the data into
dir_with_files = simpledialog.askstring(title = 'Folder path prompt', prompt = "Please type in the full path of the folder containing your files:    ")

def Merge_RPT_files():
    global main_file_check; global main_data; global main_sheet; global workbook; global worksheet; global row_count; global column_count
    extra_save = 0 # For saving a backup automatically 

    # Setting up the main file everything will go into
    if main_file_check == 0:
        main_data = pyxl.load_workbook(main_file)
        main_sheet = main_data.worksheets[0]

        # Setting up the active workbook
        global workbook
        workbook = Workbook()
        worksheet = workbook.active
        row_count = main_sheet.max_row
        column_count = main_sheet.max_column

    if stop == 0:
        stop_program()

    # Make sure data in the main file stays (because this will overwite existing file)
    for i in range (1, row_count + 1):
        for j in range (1, column_count + 1):
            # reading cell value from source file
            c = main_sheet.cell(row = i, column = j)
            # writing the read value to destination file
            worksheet.cell(row = i, column = j).value = c.value
            # worksheet.cell(row = i, column = j).fill  = c.fill # keep any background filling
            main_file_check = 1

    if stop == 0:
        stop_program()

    for file in os.listdir(dir_with_files):
        os.chdir(dir_with_files)
        book = pyxl.load_workbook(file, data_only = True)
        names = book.sheetnames
        # Have to get index for specific sheet
        i = 0
        for name in names:
            if name == 'INSP RPT 2':
                break
            else:
                i += 1
        sheet = book.worksheets[i]

        # Move data into the main workbook
        # Columns to copy are A10 - AA57
        for row in range(10, 57+1): # Rows 10-57
            for column in range(1, 27+1): # Columns A-AA
                c = sheet.cell(row, column)
                worksheet.cell(row = row_count+1, column = column).value = c.value
                column += 1
            row += 1; row_count += 1 
        
        os.remove(file)
        
        if stop == 0:
            stop_program()
    
    finishing_touches()
    workbook.save(filename = main_file)
    text = Label(app, text = "Finished merging \n Saving, one moment")
    text.grid()
    stop_program()


def finishing_touches():
    global main_file_check; global main_data; global main_sheet; global workbook; global worksheet; global row_count; global column_count
    # Setting up the main file everything will go into
    if main_file_check == 0:
        main_data = pyxl.load_workbook(main_file)
        main_sheet = main_data.worksheets[0]

        # Setting up the active workbook
        global workbook
        workbook = Workbook()
        worksheet = workbook.active
        row_count = main_sheet.max_row
        column_count = main_sheet.max_column

    # Copy the heat number
    for row in range(5, row_count+1):
        c =  worksheet.cell(row-1,   2)
        c2 = worksheet.cell(row, 2)
        if c2.value == None:
            c2.value = c.value

    # Delete Empty Rows
    rows_to_delete = []
    for row in range(4, row_count+1): # We must first find each row that needs to be deleted
        empty = []
        for column in range(4, 22+1): # columns D-V
            c = worksheet.cell(row, column)
            if c.value == None:
                empty.append(column)
        if len(empty) == 19: # 19 is all rows in range
            rows_to_delete.append(row)
    for row in reversed(rows_to_delete): # Delete the rows >:)
        worksheet.delete_rows(row)

    ### Formatting
    # Merge cells
    worksheet.merge_cells('A1:A3'); worksheet.merge_cells('B1:B3'); worksheet.merge_cells('C1:C3')
    worksheet.merge_cells('D1:G1'); worksheet.merge_cells('D2:E2'); worksheet.merge_cells('F2:G2')
    worksheet.merge_cells('H1:S1'); worksheet.merge_cells('H2:M2'); worksheet.merge_cells('N2:P2'); worksheet.merge_cells('Q2:S2')
    worksheet.merge_cells('T1:T3'); worksheet.merge_cells('U1:U3'); worksheet.merge_cells('V1:V3'); worksheet.merge_cells('W1:Y3'); worksheet.merge_cells('X1:X3'); worksheet.merge_cells('Y1:Y3'); worksheet.merge_cells('Z1:Z3'); worksheet.merge_cells('AA1:AA3')
    for i in range(4, row_count+1):
        worksheet.merge_cells('W{}:Y{}'.format(i,i))

    # Center every cell
    for row in worksheet:
        for cell in row:
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')

    # Bold the font on the Header
    for row in range(1,3+1):
        for col in range(1, column_count):
            worksheet.cell(row = row, column = col).font = Font(bold=True)

    ### Apply borders
    # Put a thin border on every cell in the worksheet
    thin_border = Border(left=Side(style='thin'), 
                            right=Side(style='thin'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin'))
    for i in range(1,row_count+1):
        for j in range(1, column_count):
            worksheet.cell(row = i, column = j).border = thin_border

    # Put a thick right border of select columns
    columns_for_border = [3, 7, 13, 19, 27]
    thick_border = Border(left=Side(style='thin'), 
                            right=Side(style='thick'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin'))
    for col in columns_for_border:
        for i in range(1, row_count+1):
            worksheet.cell(row = i, column = col).border = thick_border

    # Add some special header borders
    columns_for_border = [2, 21, 22, 23, 24, 25, 26]
    bottom_header_border = Border(left=Side(style='thin'), 
                                    right=Side(style='thin'), 
                                    top=Side(style='thick'), 
                                    bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = bottom_header_border

    columns_for_border = [5, 6, 9, 10, 11, 12, 15, 16, 17, 18]
    mid_header_border = Border(left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = mid_header_border

    columns_for_border = [1, 20]
    left_header_border = Border(left=Side(style='thick'), 
                                right=Side(style='thin'), 
                                top=Side(style='thick'), 
                                bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = left_header_border

    columns_for_border = [3, 27]
    right_header_border = Border(left=Side(style='thin'), 
                                    right=Side(style='thick'), 
                                    top=Side(style='thick'), 
                                    bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = right_header_border

    columns_for_border = [4, 8, 14]
    mid_left_header_border = Border(left=Side(style='thick'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = mid_left_header_border

    columns_for_border = [7, 13, 19]
    mid_right_header_border = Border(left=Side(style='thin'), 
                                right=Side(style='thick'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thick'))
    for col in columns_for_border:
        worksheet.cell(row = 3, column = col).border = mid_right_header_border



def start_thread():
    # For repetition
    global main_file_check
    main_file_check = 0

    # Assign global variable and initialize value
    global stop
    stop = 1
    text = Label(app, text = 'Merging')
    text.grid()

    # Create and launch a thread 
    t = Thread (target = Merge_RPT_files)
    t.start()

def stop():
    # Assign global variable and set value to stop
    text = Label(app, text = 'Please wait, do not close the program yet')
    text.grid()
    global stop
    stop = 0

def stop_program():
    # Put on the finishing touches
    finishing_touches()
    # Save the excel file and close the workbooks
    try:
        workbook.save(filename = main_file)
    except:
        pass
    try: # Just in case the backup is open for any reason
        os.chdir(dir_with_files)
        # shutil.copy(main_file, 'RPT Backup.xlsx')
    except:
        pass

    text = Label(app, text = "The program may be closed now")
    text.grid()
    sys.exit()

root_win = tk.Tk()
root_win.title("RPT Merge")
root_win.geometry('400x300')
app = Frame(root_win)
app.grid()
start_button = Button(app, text="Start the Merge",command=start_thread)
stop_button = Button(app, text="Stop Merging",command=stop)

start_button.grid()
stop_button.grid()

app.mainloop()
