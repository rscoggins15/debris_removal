'''
This project pulls ticket data from a spreadsheet and generates reports that summarize 
quantities for specific roadways and inspectors. This generator allows for easy compilation 
of report data to be pasted over into SiteManager for data entry.
'''

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import askopenfilename

import pandas as pd


# generate the report
def generate_reports(df, output_file_path, selection):
    '''Opens output file, pulls data together, and writes it to the report file.'''
    with open(output_file_path, 'w', encoding='utf-8') as file:
        # Group the data by Date
        selected_date = df[df['Date'] == selection]
        grouped_by_date = selected_date.groupby('Date')

        for date, date_group in grouped_by_date:

            # Group the data by Loading Site Monitor (Inspector) for each date
            grouped_by_inspector = date_group.groupby('Loading Site Monitor')

            for inspector, inspector_group in grouped_by_inspector:
                # Group the data by Roadway for each inspector on that date
                grouped_by_roadway = inspector_group.groupby('Roadway')

                for roadway, roadway_group in grouped_by_roadway:
                    # Group the data by TRM for each roadway for that inspector on that date
                    grouped_by_trm = roadway_group.groupby('TRM')

                    for trm, trm_group in grouped_by_trm:
                        # Extracting unique values for each entry in the report
                        county = trm_group['County'].unique()[0]
                        disposal_site = trm_group['Disposal Site'].unique()[0]

                        # Calculate the quantities and ticket counts for each type of debris
                        gen_debris_removal_hwy_row = trm_group[trm_group['Material'] == 'Gen Debris Removal HWY ROW']['Net Quantity'].sum()
                        gen_debris_ticket_count = trm_group[trm_group['Material'] == 'Gen Debris Removal HWY ROW'].shape[0]
                        
                        leaning_trees_ea_tree = trm_group[trm_group['Material'] == 'Leaning Trees EA Tree']['Net Quantity'].sum()
                        leaning_trees_ticket_count = trm_group[trm_group['Material'] == 'Leaning Trees EA Tree'].shape[0]
                        
                        hanging_limbs_ea_tree = trm_group[trm_group['Material'] == 'Hanging Limbs EA Tree']['Net Quantity'].sum()
                        hanging_limbs_ticket_count = trm_group[trm_group['Material'] == 'Hanging Limbs EA Tree'].shape[0]

                        stumps = trm_group[trm_group['Material'] == 'Tree Stump Removal']['Net Quantity'].sum()
                        stumps_ticket_count = trm_group[trm_group['Material'] == 'Tree Stump Removal'].shape[0]

                        # Write the report for this unique combination to the file
                        file.write(f"Date: {date}\n")
                        file.write(f"Data Entry for: {inspector}\n")
                        file.write(f"County: {county}\n")
                        file.write(f"Roadway: {roadway}\n")
                        file.write(f"TRM range: {trm}\n")
                        file.write(f"Disposal Site: {disposal_site}\n")
                        file.write(f"Gen Debris Removal HWY ROW: {round(gen_debris_removal_hwy_row, 3)}  [{gen_debris_ticket_count}]\n") if gen_debris_ticket_count > 0 else 0
                        file.write(f"Leaning Trees EA Tree: {leaning_trees_ea_tree}  [{leaning_trees_ticket_count}]\n") if leaning_trees_ticket_count > 0 else 0
                        file.write(f"Hanging Limbs EA Tree: {hanging_limbs_ea_tree}  [{hanging_limbs_ticket_count}]\n") if hanging_limbs_ticket_count > 0 else 0
                        file.write(f"Stumps EA: {stumps}  [{stumps_ticket_count}]\n") if stumps_ticket_count > 0 else 0
                        file.write("\n")


def report_folder():
    '''
    Checks if a folder exists. If it does not, it creates one
    '''
    # global report_folder_path
    # current_dir = os.path.dirname(os.path.abspath(__file__))
    report_folder_path = 'reports'
    if not os.path.exists(report_folder_path):
        try:
            os.mkdir(report_folder_path)
            print(f'Reports folder created: {report_folder_path}')
        except Exception as e:
            messagebox.showerror("Error Occured", e)
    # else:
        # print(f'Reports folder already exists: {report_folder_path}')
    return report_folder_path

# function to generate report
def report_gen():
    global selection
    selection = combo.get()

    # Define the output file path
    selected_date = selection.replace("/", "-")
    path = excel_file_path.split('/')
    county = str(path[-2])
    output_file_path = os.path.join(report_folder(), f'report_{county}_{selected_date}.txt')

    # Call the function to generate the reports and write to the file
    try:
        generate_reports(df, output_file_path, selection)
        # Show message box that file has been written to
        confimation_label_show()
        # messagebox.showinfo('Report Generation', "Report generated successfully!")

    except Exception as err:
        # handle the error
        messagebox.showerror("Error Occured:", err)   # printing exception that occured

def confirmation_label_hide():
    '''
    This will hide the confirmation label
    '''
    confirm.config(text='')

def confimation_label_show():
    '''
    This will show a success confirmation upon sueccessful report generation
    for 5 seconds and then call another function to make it disappear. 
    '''
    confirm.config(text='Report generated successfully!')
    root.after(5000, confirmation_label_hide)


def ask_for_file():
    '''Popup to get a file name and dir.'''
    global excel_file_path
    excel_file_path = askopenfilename(initialdir = 'T:\\TylMaintenance\\Contracts\\Emergency Contracts\\FY 24 EMERGENCY DEBRIS REMOVAL\\TXDOT Daily Inspection Report')
    v.set(os.path.basename(excel_file_path))

def restart():
    '''
    restarts the program, so a new file can be selected
    '''
    os.execv(sys.executable, ['python'] + sys.argv)

# GUI
# Define window
root = tk.Tk()
root.title("Daily Report")
v = tk.StringVar()

'''
Reads file from user
Will repeat this until user selects a valid spreadsheet file
'''
while True:
    try:
        ask_for_file()
        # read excel file into df
        df = pd.read_excel(excel_file_path, sheet_name='Ticket List')

        # define columns of interest
        columns_of_interest = ['Date', 'Loading Site Monitor', 'Roadway', 'TRM', 'County', 'Disposal Site', 'Material', 'Net Quantity']

        # Filter df to only include the columns of interest
        df = df[columns_of_interest]

        unique_dates = df['Date'].dropna().unique().tolist()
        break

    except Exception as error:
        # handle the error
        answer = messagebox.askyesno("Error Occured:", f'{error}\n\n\nCONTINUE?')   # printing exception that occured
        if answer:
            continue # If user clicks yes, continue
        sys.exit() # If user clicks no, quit code

# Change file
btn_change_date = ttk.Button(root, text = "Select a New File", command = restart)
btn_change_date.grid(column = 0, row = 0, padx = 2, pady = 2)

# padding
ttk.Label(text= ' ').grid(column = 0, row = 1)

# Display selected file
chosen_file_text = ttk.Label(text='Chosen File:')
chosen_file_text.grid(column = 0, row = 2, sticky = 'E')
chosen_file = ttk.Label(textvariable = v, font = 'Helvetica 8 bold')
chosen_file.grid(column = 1, row = 2)

# Prompt
instrucitons = ttk.Label(text = "Select a Date:")
instrucitons.grid(column = 0, row = 3, sticky = 'E')

# Combobox
combo = ttk.Combobox(root, values = unique_dates)
combo.grid(column = 1, row = 3, sticky = 'W')
combo.current(0)

# Confirmation text
confirm = ttk.Label(root, text='', font=('Helvetica',10, 'bold'))
confirm.grid(column=1, row=4, sticky='W', pady='10')

# Buttons
generate_btn = ttk.Button(root, text = "Generate Report", command = report_gen)
generate_btn.grid(column = 2, row = 3, padx = 2, sticky = 'W')

#padding
ttk.Label(text = ' ').grid(column = 0, row = 4)

root.mainloop()
