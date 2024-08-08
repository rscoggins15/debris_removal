import pandas as pd
import os
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter as tk


# generate the report
def generate_reports(df, output_file_path, selection):
    with open(output_file_path, 'w') as file:
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



# fuction to generate report
def report_gen():
    global selection
    selection = combo.get()

    # Define the output file path
    selected_date = selection.replace("/", "-")
    path = excel_file_path.split('/')
    county = str(path[-2])
    output_file_path = os.path.join(os.path.dirname(__file__), f'report_{county}_{selected_date}.txt')

    # Call the function to generate the reports and write to the file
    generate_reports(df, output_file_path, selection)


def ask_for_file():
    global excel_file_path
    excel_file_path = askopenfilename(initialdir = 'T:\\TylMaintenance\\Contracts\\Emergency Contracts\\FY 24 EMERGENCY DEBRIS REMOVAL\\TXDOT Daily Inspection Report')
    v.set(os.path.basename(excel_file_path))

# get file name
def file_name():
    return excel_file_path




# GUI

# Define window
root = tk.Tk()
root.title("Daily Report")
v = tk.StringVar()

ask_for_file()


# read excel file into fd
df = pd.read_excel(excel_file_path, sheet_name='Ticket List')

# define columns of interest
columns_of_interest = ['Date', 'Loading Site Monitor', 'Roadway', 'TRM', 'County', 'Disposal Site', 'Material', 'Net Quantity']

# Filter df to only include the columns of interest
df = df[columns_of_interest]

unique_dates = df['Date'].dropna().unique().tolist()


# Change file
# btn_change_date = ttk.Button(root, text = "Select a New File", command = ask_for_file)
# btn_change_date.grid(column = 0, row = 0, padx = 2, pady = 2)

# padding
lbl_padding = ttk.Label(text= ' ').grid(column = 0, row = 1)

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

# Buttons
generate_btn = ttk.Button(root, text = "Generate Report", command = report_gen)
generate_btn.grid(column = 2, row = 3, padx = 2, sticky = 'W')

#padding
lbl_padding2 = ttk.Label(text = ' ').grid(column = 0, row = 4)


root.mainloop()