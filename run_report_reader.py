import tabula
import pandas
import pathlib
import numpy as np
import re
from collections import OrderedDict
import csv
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import os

#########################################
# This block of code creates the GUI.
# These functions define what happens when you click the various buttons

open_filepath = None
save_filepath = None

def open_file():
    global open_filepath
    open_filepath = askopenfilename(filetypes=[('PDF', '*.pdf')])
    open_var.set(open_filepath)
    if not open_filepath:
        return

def save_file():
    global save_filepath
    save_filepath = asksaveasfilename(defaultextension='csv', filetypes=[("CSV", "*.csv")])
    save_var.set(save_filepath)
    if not save_filepath:
        return

# This function defines a message box that displays the appropriate error message
# when called with the error message and window title as arguments.
# We use this to alert the user if their input is invalid
def error_box(box_title, box_label):
    error_window = tk.Tk()
    error_window.title(box_title)
    error_window.geometry("500x100")
    error_lbl = tk.Label(master=error_window, text=box_label)
    error_lbl.pack()
    error_window.mainloop()

# This function defines what happens when you click the Next button.  The variables
# are all designated as global to allow the program to use these values outside of the
# window.  Possible future upgrade would to use an object oriented approach here
# rather than the global variables.
def take_ent():
    # this is where we check that the source and save file paths have been provided.

    if open_filepath == None or open_filepath == "":
        error_box("Enter All Fields", "Please select a source file.")

    elif save_filepath == None or save_filepath == "":
        error_box("Enter All Fields", "Please select a save location.")

    else:
        window.destroy()

window = tk.Tk()
window.title("Load Step Chart Maker")
window.geometry("900x600")

#open source file location
frm_1 = tk.Frame(master = window)
frm_1.grid(row=1, column=0, padx=20, pady=(20,10), sticky='w')
btn_open = tk.Button(master = frm_1, text = "Select source file", width=20, command=open_file)
btn_open.grid(row=0, column=0)

open_var = tk.StringVar()
lbl_open = tk.Label(master=frm_1, textvariable=open_var, width=70, relief=tk.SUNKEN, borderwidth=2)
lbl_open.grid(row=0, column=1, padx=20)

#save result file location
frm_2 = tk.Frame(master=window)
frm_2.grid(row = 12, column = 0, padx = 20, pady = 10, sticky = 'w')
btn_save = tk.Button(master = frm_2, text='Results save location', width=20, command=save_file)
btn_save.grid(row=0, column=0)

save_var = tk.StringVar()
lbl_save = tk.Label(master=frm_2, textvariable=save_var, width=70, relief=tk.SUNKEN, borderwidth=2)
lbl_save.grid(row=0, column=1, padx=20)

# Next Button
btn = tk.Button(master=window, text="Next", width = 15, height = 3, command = take_ent)
btn.grid(row=15, column=1, sticky='se')


window.mainloop()

# End of GUI portion
######################################################################

# pdf_path_1 = r"C:\Users\Jay\Desktop\Python\Auto Run File Builder\test_run_file.pdf"
# pdf_path_2 = r"C:\Users\Jay\Desktop\Python\Auto Run File Builder\test_run_file.pdf"
# pdf_path_3 = r"C:\Users\Jay\Desktop\Python\Auto Run File Builder\targa_kbz.pdf"
# csv_path = r"C:\Users\Jay\Desktop\Python\Auto Run File Builder\test.csv"
# tabula.convert_into(pdf_path, csv_path, output_format="csv", stream=True, pages="all")

df = tabula.read_pdf(open_filepath, pages="all", multiple_tables=True)

def fix_column_headers(list_of_dfs):
    n = 0
    for frame in list_of_dfs:
        col_names_default = frame.columns
        new_row={}
        col_num = []
        for i in range(len(frame.columns)):
            col_num.append(i)
            new_row[i] = col_names_default[i]
        new_row_df = pandas.DataFrame(new_row, index=[0])
        frame.columns = col_num
        list_of_dfs[n] = pandas.concat([new_row_df, frame]).reset_index(drop=True)
        n+=1

def create_df_dict(list_of_dfs):
    for frame in list_of_dfs:
        for index, row in frame.iterrows():
            combined_df_dict[row[0]] = [row[i] for i in range(1, len(frame.columns))]

fix_column_headers(df)

combined_df_dict = {}
create_df_dict(df)

def top_section_corrector(word):
    for key, value in combined_df_dict.items():
        if key.startswith(word):
            word_line = [key]
            word_line.extend(value)
            word_line = [str(item) for item in word_line if str(item) != 'nan']
            word_line_joined = ' '.join(word_line)
            return word_line_joined

elevation_line_joined = top_section_corrector('Elevation')
frame_line_joined = top_section_corrector('Frame')
max_rl_line_joined = top_section_corrector('Max RL Tot')
rated_rpm_line_joined = top_section_corrector('Rated RPM')
calc_rpm_line_joined = top_section_corrector('Calc RPM')
#######################
#######################
#create separate key value paires for each value in the
# top section of the report

#Line 1 Elevation
combined_df_dict['Elevation,ft']=re.search(r'Elevation,ft:(.*?)Barmtr', elevation_line_joined).group(1).strip()
combined_df_dict['Barmtr,psia'] = re.search(r'Barmtr,psia:(.*?)Ambient', elevation_line_joined).group(1).strip()
combined_df_dict['Ambient,F'] = re.search(r'Ambient,F:(.*?)Type', elevation_line_joined).group(1).strip()
combined_df_dict['Driver Type'] = re.search(r'Type:(.*)', elevation_line_joined).group(1).strip()

#Line 2 Frame
combined_df_dict['Frame'] = re.search(r'Frame:(.*?)Stroke', frame_line_joined).group(1).strip()
combined_df_dict['Stroke'] = re.search(r'Stroke, in:(.*?)Rod', frame_line_joined).group(1).strip()
combined_df_dict['Rod Dia'] = re.search(r'Dia, in:(.*?)Mfg', frame_line_joined).group(1).strip()
if re.search(r'Mfg:(.*)', frame_line_joined).group(1):
    combined_df_dict['Driver Mfg'] = re.search(r'Mfg:(.*)', frame_line_joined).group(1)
else:
    combined_df_dict['Driver Mfg'] = "none"

#Line 3 Max RL
combined_df_dict['Max RL Tot'] = re.search(r'Max RL Tot, lbf:(.*?)Max RL Tens', max_rl_line_joined).group(1).strip()
combined_df_dict['Max RL Tens'] = re.search(r'Max RL Tens, lbf:(.*?)Max RL Comp', max_rl_line_joined).group(1).strip()
combined_df_dict['Max RL Comp'] = re.search(r'Max RL Comp, lbf:(.*?)Model', max_rl_line_joined).group(1).strip()
if re.search(r'Model:(.*)', max_rl_line_joined).group(1):
    combined_df_dict['Driver Model'] = re.search(r'Model:(.*)', max_rl_line_joined).group(1)
else:
    combined_df_dict['Driver Model'] = 'none'

#Line 4 Rated RPM
combined_df_dict['Rated RPM'] = re.search(r'Rated RPM:(.*?)Rated BHP', rated_rpm_line_joined).group(1).strip()
combined_df_dict['Rated BHP'] = re.search(r'Rated BHP:(.*?)Rated PS', rated_rpm_line_joined).group(1).strip()
combined_df_dict['Rated PS'] = re.search(r'Rated PS FPM:(.*?)BHP', rated_rpm_line_joined).group(1).strip()
combined_df_dict['BHP'] = re.search(r'\d BHP:(.*)', rated_rpm_line_joined).group(1).strip()

#Line 5 Calc RPM
combined_df_dict['Calc RPM'] = re.search(r'Calc RPM:(.*?)BHP:', calc_rpm_line_joined).group(1).strip()
combined_df_dict['Calc BHP'] = re.search(r'BHP:(.*?)Calc PS', calc_rpm_line_joined).group(1).strip()
combined_df_dict['Calc PS'] = re.search(r'Calc PS FPM:(.*?)Avail', calc_rpm_line_joined).group(1).strip()
combined_df_dict['Avail HP'] = re.search(r'Avail:(.*)', calc_rpm_line_joined).group(1).strip()


# export that dict to JSON and .csv


# program currently only takes either the Flow Calc, MMSCFD from the
# Stage data or the cylinder data, not both and behavior depends on
# how tabula breaks up the data into dataframes and which gets
# added to the combined_df_dict second (overwriting the line added
# first)

ordered_key_list = [
    'Elevation,ft',
    'Barmtr,psia',
    'Ambient,F',
    'Driver Type',
    'Frame',
    'Stroke',
    'Rod Dia',
    'Driver Mfg',
    'Max RL Tot',
    'Max RL Tens',
    'Max RL Comp',
    'Driver Model',
    'Rated RPM',
    'Rated BHP',
    'Rated PS',
    'BHP',
    'Calc RPM',
    'Calc BHP',
    'Calc PS',
    'Avail HP',
    'Services',
    'Gas Model',
    'Stage Data:',
    'Target Flow, MMSCFD',
    'Flow Calc, MMSCFD',
    'BHP per Stage',
    'Specific Gravity',
    'Ratio of Sp Ht (N)',
    'Comp Suct (Zs)',
    'Comp Disch (Zd)',
    'Pres Suct Line, psig',
    'Pres Suct Flg, psig',
    'Pres Disch Flg, psig',
    'Pres Disch Line, psig',
    'Pres Ratio F/F',
    'Temp Suct, F',
    'Temp Clr Disch, F',
    'Cylinder Data:',
    'Cyl Model',
    'Cyl Bore, in',
    'Cyl RDP (API), psig',
    'Cyl MAWP, psig',
    'Cyl Action',
    'Cyl Disp, CFM',
    'Pres Suct Intl, psig',
    'Temp Suct Intl, F',
    'Pres Disch Intl, psig',
    'Temp Disch Intl, F',
    'HE Suct Gas Vel, FPM',
    'HE Disch Gas Vel, FPM',
    'HE Spcrs Used/Max',
    'HE Vol Pkt Avail',
    'Vol Pkt Used',
    'HE Min Clr, %',
    'HE Total Clr, %',
    'CE Suct Gas Vel, FPM',
    'CE Disch Gas Vel, FPM',
    'CE Spcrs Used/Max',
    'CE Min Clr, %',
    'CE Total Clr, %',
    'Suct Vol Eff HE/CE, %',
    'Disch Event HE/CE, ms',
    'Suct Pseudo-Q HE/CE',
    'Gas Rod Ld Comp, %',
    'Gas Rod Ld Tens, %',
    'Gas Rod Ld Total, %',
    'Xhd Pin Deg/%Rvrsl lbf',
    'Cyl BHP'
]

combined_ordered_dict = OrderedDict()
for key in ordered_key_list:
    for dict_key in combined_df_dict.keys():
        if dict_key.startswith(key):
            combined_ordered_dict[key] = combined_df_dict[dict_key]

# print(type(combined_ordered_dict['Cyl BHP']))

for key, value in combined_ordered_dict.items():
    if type(value) is list:
        combined_ordered_dict[key] = [item for item in value if str(item) != 'nan']

with open(pathlib.Path(save_filepath), 'w') as csvfile:
    writer = csv.writer(csvfile)
    for key, value in combined_ordered_dict.items():
        writer.writerow([key, value])

os.startfile(save_filepath)
