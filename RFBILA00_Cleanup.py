import numpy as np
import pandas as pd
from pandas import DataFrame
import re
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox

# Define global vars
import_file = ""
import_filepath = ""
import_filename = ""
export_filename = ""

# Fuctions Definitions
def UploadAction(event=None):
    """Browse for excel file button function
    Function will open a browse window to select an excel file.
    GUI label texts will change appropriately.
    """
    global import_file, import_filepath, import_filename, export_filename
    import_file = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls")])
    import_filepath = os.path.split(import_file)[0]
    import_filename = os.path.split(import_file)[1]
    export_filename = import_filename.split('.')[0]+str('_cleanup')+str('.xlsx')
    
    selected_label['text'] = "Selected file: " + str(import_filename)
    output_label['text'] = "The tool will create a new Excel file\nnamed:\n"+str(export_filename)+"\nin the same folder as the imported file."
    statusbar['text'] = "Ready to import the file."

def UpdateScrollRegion(event):
    """Tkinter Scroll Region Update
    This fuction is used to update the scroll region when previewing imported excel file.
    """
    input_canvas.configure(scrollregion=input_canvas.bbox('all'))

def ImportXLS(event=None):
    """Import XLS file button function
    Function will import the selected excel file and display a preview
    of the file (first 15 rows including header) in the lower frame.
    """
    global xls
    if import_file != "":
        xls = pd.read_excel(import_file, sheet_name = 0, header = 0)
        
        scrollbar = tk.Scrollbar(input_canvas, orient='horizontal', command=input_canvas.xview)
        scrollbar.pack(side="bottom", fill='x')

        input_canvas.configure(xscrollcommand = scrollbar.set)

        import_frame = tk.Frame(input_canvas, bg="white")
        input_canvas.create_window((0,0), window=import_frame, anchor='nw')
        input_canvas.bind('<Configure>', UpdateScrollRegion)
        import_frame.bind('<Configure>', UpdateScrollRegion)

        for idx, x in enumerate(xls.columns):
            tk.Label(import_frame, text="Column "+str(idx), bg='white', font=bold_font).grid(row=0, column=idx, sticky='W')
            tk.Label(import_frame, text=x, bg='white').grid(row=1, column=idx, sticky='W')

        for x in range(xls.head(13).shape[1]):
            for y in range(13):
                tk.Label(import_frame, text=xls.iat[y,x], bg='white').grid(row=y+2, column=x, sticky='W')
                
        statusbar['text'] = "RFBILA00 Excel file imported. Please set up the parameters and generate a cleaned up version."
    else:
        statusbar['text'] = "No file selected. Please select an RFBILA00 Excel file."

def regexcontains(s):
    return re.search('[a-zA-Z0-9]', s)

def RFBILACleanup(xls, xls_hierarchy, xls_account_num, xls_account_name, xls_account_val_current, xls_account_val_previous, xls_account_pos, xls_account_sum):
    '''Cleans up RFBIL00 excel sheets and generates a new excel file
    The clean up consists of placing account hierarchies in the same row as the corresponding accounts.
    Features: identification of multi-line hierarchy names and identification of account hierarchy sums.
    '''
    
    # Add working column into DataFrame 
    xls['##Category##'] = 'Blank'
    xls_num_columns = xls.shape[1]

    # Identify row type and set row type information into ##Category## column
    xls_copy = xls.copy() # Dataframe copy because pandas doues not take into account changes in the dataframe during iterations
    for idx, row in xls.iterrows():
        if not pd.isnull(row[xls_account_num]):
            xls.at[idx,'##Category##'] = 'Account'
        else:
            if pd.isnull(row[xls_account_name]):
                xls.at[idx,'##Category##'] = 'Delete'
            elif regexcontains(row[xls_account_name]) is None:
                xls.at[idx,'##Category##'] = 'Delete'
            elif (row[xls_account_val_current] + row[xls_account_val_previous]) < 0 or (row[xls_account_val_current] + row[xls_account_val_previous]) > 0:
                xls.at[idx,'##Category##'] = 'Hierarchy Sum'
                # Check if sum identificator is not multi-row
                if xls.iat[idx+1, xls_hierarchy] == xls.iat[idx, xls_hierarchy]:
                    if regexcontains(xls.iat[idx+1, xls_account_name]) is not None:
                        if xls.iat[idx+1, xls_account_pos] == xls.iat[idx, xls_account_pos]:
                            if pd.isnull(xls.iat[idx+1, xls_account_num]):
                                xls.iat[idx+1, xls_account_name] = str(xls_account_sum) + str(xls.iat[idx+1, xls_account_name])
                                xls_copy = xls.copy()
            # Special consideration for sum of hierarchies
            else:
                try:
                    # Check if account name starts with sum identificator
                    if re.sub(' +', ' ',xls_copy.iat[idx, xls_account_name].strip()).startswith(str(xls_account_sum)):
                        xls.at[idx,'##Category##'] = 'Hierarchy Sum'
                        # Check if sum identificator is not multi-row
                        if xls.iat[idx+1, xls_hierarchy] == xls.iat[idx, xls_hierarchy]:
                            if regexcontains(xls.iat[idx+1, xls_account_name]) is not None:
                                if xls.iat[idx+1, xls_account_pos] == xls.iat[idx, xls_account_pos]:
                                    if pd.isnull(xls.iat[idx+1, xls_account_num]):
                                        xls.iat[idx+1, xls_account_name] = str(xls_account_sum) + str(xls.iat[idx+1, xls_account_name])
                                        xls_copy = xls.copy()
                except:
                    pass
    del xls_copy


    # Clean multiline Hierarchies
    for idx, row in xls[::-1].iterrows():
        if idx > 0:
            if row['##Category##'] == 'Blank' and \
            (xls.at[idx-1,'##Category##'] =='Blank' or xls.at[idx-1,'##Category##'] == 'Hierarchy') and \
            row[xls_hierarchy] == xls.iat[idx-1,xls_hierarchy]:
                xls.at[idx-1,'##Category##'] = 'Hierarchy'
                new_name = re.sub(' +', ' ',xls.iat[idx-1, xls_account_name].strip()) + ' ' + re.sub(' +', ' ',xls.iat[idx, xls_account_name].strip())
                xls.iat[idx-1, xls_account_name] = new_name
                xls.at[idx,'##Category##'] = 'Delete'

    # Deal with the rest Hierarchies item in dataframe
    for idx, row in xls.iterrows():
        if xls.at[idx,'##Category##'] == 'Blank':
            if regexcontains(xls.iat[idx,xls_account_name]) is not None and ((xls.iat[idx,xls_account_val_current] + xls.iat[idx,xls_account_val_previous] == 0) or (pd.isnull(xls.iat[idx,xls_account_val_current]) and pd.isnull(xls.iat[idx,xls_account_val_previous]))) and pd.isnull(xls.iat[idx,xls_account_num]):
                xls.at[idx,'##Category##'] = 'Hierarchy'
                xls.iat[idx, xls_account_name] = re.sub(' +', ' ',xls.iat[idx, xls_account_name].strip())

    # Create Hierarchy Columns
    hierarchy_count = xls[xls.columns[xls_hierarchy]].max() - xls[xls.columns[xls_hierarchy]].min()
    hierarchy_name = ""

    for i in range(hierarchy_count):
        hierarchy_name = 'Hierarchy ' + str(i)
        xls[hierarchy_name] = ''

    #Hierarchy Iterator
    hierarchy_offset = xls[xls.columns[xls_hierarchy]].min()
    current_hierarchy_true = hierarchy_offset - 1
    current_hierarchy_indicated = - 1
    hierarchy_names = hierarchy_count * ['']
    hierarchy_counters = hierarchy_count * [-1]

    for idx, row in xls.iterrows():
        # If Hierarchy
        if xls.at[idx,'##Category##'] == 'Hierarchy':
            # Case 1: Append New Hierarchy
            if xls.iat[idx, xls_hierarchy] > current_hierarchy_true:
                current_hierarchy_true = xls.iat[idx, xls_hierarchy]
                current_hierarchy_indicated += 1
                hierarchy_names[current_hierarchy_indicated] = xls.iat[idx, xls_account_name]
                hierarchy_counters[current_hierarchy_indicated] = current_hierarchy_true
            
            # Case 2: Drop Hierarchy
            elif xls.iat[idx, xls_hierarchy] <= current_hierarchy_true:
                current_hierarchy_true = xls.iat[idx, xls_hierarchy]
                for idx2, item in enumerate(hierarchy_counters):
                    if item >= current_hierarchy_true:
                        item = current_hierarchy_true
                        current_hierarchy_indicated = idx2
                        hierarchy_names[current_hierarchy_indicated] = xls.iat[idx, xls_account_name]
                        hierarchy_counters[current_hierarchy_indicated] = current_hierarchy_true
                        hierarchy_names[(idx2+1):] = (len(hierarchy_names)-idx2 -1) * ['']
                        hierarchy_counters[(idx2+1):] = (len(hierarchy_counters)-idx2 -1) * [-1]
                        break
        
        # If Account
        if xls.at[idx,'##Category##'] == 'Account':
            for idx3, item in enumerate(hierarchy_names):
                xls.iat[idx, (xls_num_columns + idx3)] = hierarchy_names[idx3]

    # Output file
    options = {}
    options['strings_to_formulas'] = False
    options['strings_to_urls'] = False
    writer = pd.ExcelWriter(os.path.join(import_filepath, export_filename), options=options)
    xls.to_excel(writer, index=False)
    writer.save()
    statusbar['text'] = "RFBILA00 excel file cleaned up successfully!"
    tk.messagebox.showinfo("Completed!", "The RFBIL00 has been successfully cleaned up.")



## GUI
root = tk.Tk()
bold_font = tkfont.Font(size=9, weight="bold")

main_canvas = tk.Canvas(root, height=625, width=885)
main_canvas.pack()

#### GUI Left Frame ####
step1_frame = tk.LabelFrame(root, text="Step 1: Import File")
step1_frame.place(height=260, width=230, x=10, y=10)

browse_label = tk.Label(step1_frame, text="Select an RFBILA00 Excel file:")
browse_label.place(relx=0.025, rely=0.01)

browse_button = tk.Button(step1_frame, text='Browse RFBILA00 Excel', command=UploadAction)
browse_button.place(relwidth=0.7, relx=0.15, rely=0.125)

excel_notice_label = tk.Label(step1_frame, justify="left", text="Notice:\nMake sure that the Excel file contains\nonly one sheet and column descriptions\nare in the first row!")
excel_notice_label.place(relx=0.025, rely=0.281)

selected_label = tk.Label(step1_frame, text="Selected file: ")
selected_label.place(relx=0.025, rely=0.625)

import_button = tk.Button(step1_frame, text='Import RFBILA00', command=ImportXLS)
import_button.place(relwidth=0.7, relx=0.15, rely=0.75)

#### GUI Center Frame ####
step2_frame = tk.LabelFrame(root, text="Step 2: Set up the excel sheet parameters")
step2_frame.place(height=260, width=390, x=250, y=10)

hierarchy_row = tk.Label(step2_frame, text="Column number for account hierarchy:")
hierarchy_row.place(relx=0.025, rely=0.01)

hierarchy_entry = tk.Entry(step2_frame)
hierarchy_entry.place(relx=0.92, rely=0.0225, relwidth=0.05)

account_number_row = tk.Label(step2_frame, text="Column number for account number:")
account_number_row.place(relx=0.025, rely=0.1175)

account_number_entry = tk.Entry(step2_frame)
account_number_entry.place(relx=0.92, rely=0.1275, relwidth=0.05)

account_name_row = tk.Label(step2_frame, text="Column number for account description:")
account_name_row.place(relx=0.025, rely=0.2225)

account_name_entry = tk.Entry(step2_frame)
account_name_entry.place(relx=0.92, rely=0.2325, relwidth=0.05)

account_value_current_row = tk.Label(step2_frame, text="Column number for account balance of the reporting period:")
account_value_current_row.place(relx=0.025, rely=0.3275)

account_value_current_entry = tk.Entry(step2_frame)
account_value_current_entry.place(relx=0.92, rely=0.3375, relwidth=0.05)

account_value_previous_row = tk.Label(step2_frame, text="Column number for account balance of the comparison period:")
account_value_previous_row.place(relx=0.025, rely=0.4325)

account_value_previous_entry = tk.Entry(step2_frame)
account_value_previous_entry.place(relx=0.92, rely=0.4425, relwidth=0.05)

account_position_row = tk.Label(step2_frame, text="Column number for account position:")
account_position_row.place(relx=0.025, rely=0.5375)

account_position_entry = tk.Entry(step2_frame)
account_position_entry.place(relx=0.92, rely=0.5475, relwidth=0.05)

account_sum_row = tk.Label(step2_frame, text="Account hierarchy sum identifier:")
account_sum_row.place(relx=0.025, rely=0.6425)

account_sum_entry = tk.Entry(step2_frame)
account_sum_entry.place(relx=0.6, rely=0.6525, relwidth=0.37)

account_sum_notice = tk.Label(step2_frame, justify="left", text="Notice: Depending on the SAP language setting, the sum identifier is\n'Sum ' or 'Summe '. Please type in the identifier without quotation\nmarks but with the space at the end.")
account_sum_notice.place(relx=0.025, rely=0.7575)

#### GUI Right Frame #####
step3_frame = tk.LabelFrame(root, text="Step 3: Generate Cleanup") 
step3_frame.place(height=260, width=230, x=650, y=10)

output_label = tk.Label(step3_frame, justify="left", text="The tool will create a new Excel file\nnamed:\n\nin the same folder as the imported file.")
output_label.place(relx=0.025, rely=0.01)

import_button = tk.Button(step3_frame, text='Generate File', command=lambda: RFBILACleanup(xls, int(hierarchy_entry.get()), int(account_number_entry.get()), int(account_name_entry.get()), int(account_value_current_entry.get()), int(account_value_previous_entry.get()), int(account_position_entry.get()), str(account_sum_entry.get())))
import_button.place(relwidth=0.7, relx=0.15, rely=0.356)

about_label = tk.Label(step3_frame, justify="left", text="Version: 1.0.2019-05-24\n\nDeveloped by Martin Račko")
about_label.place(relx=0.025, rely=0.5625)

#### GUI Lower Frame
input_canvas = tk.Canvas(root, borderwidth=2, relief="sunken", bg="white")
input_canvas.place(height=340, width=870, x=10, y=280)

#### GUI Statusbar ####
statusbar = tk.Label(root, text="Select an RFBILA00 Excel file...", relief="sunken", anchor="w")
statusbar.pack(side="bottom", fill="x")

## GUI Settings and run command
root.title("RFBILA00 Excel Hierarchy Cleanup Tool")
root.resizable(False, False)
root.mainloop()