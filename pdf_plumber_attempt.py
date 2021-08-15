import pdfplumber
import re
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import os

open_filepath = r'C:\Users\Jay\Desktop\Python\Auto Run File Builder\test_run_file.PDF'
save_filepath = r'C:\Users\Jay\Desktop\Python\Auto Run File Builder\test_run_file.runM'

# open_filepath = None
# save_filepath = None

# def open_file():
#     global open_filepath
#     open_filepath = askopenfilename(filetypes=[('PDF', '*.pdf')])
#     open_var.set(open_filepath)
#     if not open_filepath:
#         return
#
# def save_file():
#     global save_filepath
#     save_filepath = asksaveasfilename(defaultextension='runM', filetypes=[("runM", "*.runM")])
#     save_var.set(save_filepath)
#     if not save_filepath:
#         return
#
# # This function defines a message box that displays the appropriate error message
# # when called with the error message and window title as arguments.
# # We use this to alert the user if their input is invalid
# def error_box(box_title, box_label):
#     error_window = tk.Tk()
#     error_window.title(box_title)
#     error_window.geometry("500x100")
#     error_lbl = tk.Label(master=error_window, text=box_label)
#     error_lbl.pack()
#     error_window.mainloop()
#
# # This function defines what happens when you click the Next button.  The variables
# # are all designated as global to allow the program to use these values outside of the
# # window.  Possible future upgrade would to use an object oriented approach here
# # rather than the global variables.
# def take_ent():
#     # this is where we check that the source and save file paths have been provided.
#
#     if open_filepath == None or open_filepath == "":
#         error_box("Enter All Fields", "Please select a source file.")
#
#     elif save_filepath == None or save_filepath == "":
#         error_box("Enter All Fields", "Please select a save location.")
#
#     else:
#         window.destroy()
#
# window = tk.Tk()
# window.title("Auto Run File Builder")
# window.geometry("900x600")
#
# #open source file location
# frm_1 = tk.Frame(master = window)
# frm_1.grid(row=0, column=0, padx=20, pady=(100,20), sticky='w')
# btn_open = tk.Button(master = frm_1, text = "Select source file", width=20, command=open_file)
# btn_open.grid(row=0, column=0)
#
# open_var = tk.StringVar()
# lbl_open = tk.Label(master=frm_1, textvariable=open_var, width=70, relief=tk.SUNKEN, borderwidth=2)
# lbl_open.grid(row=0, column=1, padx=20)
#
# #save result file location
# frm_2 = tk.Frame(master=window)
# frm_2.grid(row = 12, column = 0, padx = 20, pady = (50,50), sticky = 'w')
# btn_save = tk.Button(master = frm_2, text='Results save location', width=20, command=save_file)
# btn_save.grid(row=0, column=0)
#
# save_var = tk.StringVar()
# lbl_save = tk.Label(master=frm_2, textvariable=save_var, width=70, relief=tk.SUNKEN, borderwidth=2)
# lbl_save.grid(row=0, column=1, padx=20)
#
# # Next Button
# btn = tk.Button(master=window, text="Next", width = 15, height = 3, command = take_ent)
# btn.grid(row=15, column=1, sticky='se')
#
#
# window.mainloop()

# # End of GUI portion
# ######################################################################

# Opens the run report using the pdfplumber module, extracts the text and
#  saves it to one big string called 'text'
with pdfplumber.open(open_filepath) as pdf:
    first_page = pdf.pages[0]
    text = first_page.extract_text()

# Here we are just creating a dictionary that will later be used to associate
#  the frame model to it's correct product family
product_family = {}
product_family['JGN'] = r'JGN/JGQ'
product_family['JGQ'] = r'JGN/JGQ'
product_family['KB100'] = r'KB100'
product_family['JGM'] = r'JGM/JGP'
product_family['JGP'] = r'JGM/JGP'
product_family['JG'] = r'JG/JGA'
product_family['JGA'] = r'JG/JGA'
product_family['JGR'] = r'JGR'
product_family['KBK'] = r'KBK/KBT'
product_family['KBT'] = r'KBK/KBT'
product_family['KBE'] = r'KBE'
product_family['JGJ'] = r'JGJ'
product_family['JGT'] = r'JGT'
product_family['JGC'] = r'JGC/JGD/JGF'
product_family['JGD'] = r'JGC/JGD/JGF'
product_family['JGF'] = r'JGC/JGD/JGF'
product_family['KBU'] = r'KBU/KBZ'
product_family['KBZ'] = r'KBU/KBZ'
product_family['KBB'] = r'KBB/KBV'
product_family['KBV'] = r'KBB/KBV'

# units_dict = {}

# here we search the text string for particular attributes using regular
#  expressions.  We need many of these attributes to create the output_dict
# later.  For example, the variable 'flow' here will return the flow
#  units, like MMSCFD, which we need to know which units are being utilized
#  as well as to effectively search the rest of the text string for other
#  values ie Calc Flow {flow}
flow = re.search(r'Target Flow, (\D+) \d+', text).group(1).strip()
power = re.search(f'Rated RPM: \d+ Rated (\D+): \d+', text).group(1).strip()
power_for_output = power[-2:].lower()
temperature = re.search(r'Ambient,(.*?):', text).group(1).strip()
pressure = re.search(r'Cyl MAWP, (.*?) \d+', text).group(1).strip()
if re.search(r'Max RL Tot, (.*?):', text).group(1).strip() == 'lbf':
    eng_met = 'English'
    length = 'ft'
else:
    eng_met = 'Metric'
    length = 'm'

# Here we are looking at the pressure units to determine if the pressure is in
#  gauge or absolute.  Then, if the pressure units are atm, we leave them as is
#  b/c the program doesn't assign the a or g to atm.  otherwise, we take the units
#  up to but not including the last character and set that equal to
#  pressure_for_output ie psig ends up as psi
if pressure == 'atm' or pressure[-1] == 'g':
    g_or_abs = 'Gauge'
else:
    g_or_abs = 'Absolute'

if pressure == 'atm':
    pressure_for_output = 'atm'
else:
    pressure_for_output = pressure[:-1]


# here we run through every line of the text string and create dict entries for
#  each bit of information.  For the top section of the report, we single out
#  particular values.  For the lower sections, we include the whole line in
#  each key's value.  The value is left as a long string for now, later when
#  we need a particular value, we'll split the strings into lists using .split()
#  and grab the appropriate value using either the index or slicing.
#  The regular expression here is in the parentheses (.*?) or (.*). This is
#  either saying "any character except new line (.) 0 or more (*) 0 or 1 (?)"
output_dict = {}
output_dict['Company']=re.search(r'Company:(.*?)Customer:', text).group(1).strip()
output_dict['Customer'] = re.search(r'Customer:(.*)', text).group(1).strip()
output_dict['Project'] = re.search(r'Project:(.*)', text).group(1).strip()


output_dict[f'Elevation']=re.search(fr'Elevation,{length}:(.*?)Barmtr', text).group(1).strip()
output_dict['Barmtr,psia'] = re.search(r'Barmtr,psia:(.*?)Ambient', text).group(1).strip()
output_dict['Ambient,F'] = re.search(r'Ambient,F:(.*?)Type', text).group(1).strip()
output_dict['Driver Type'] = re.search(r'Type:(.*)', text).group(1).strip()

output_dict['Frame'] = re.search(r'Frame:(.*?)Stroke', text).group(1).strip()
output_dict['Stroke'] = re.search(r'Stroke, in:(.*?)Rod', text).group(1).strip()

# Rod Dia and Model require a bit of logic since they are the only two values that
#  can be left blank.  When blank, we manually input the 'none' string as the value
output_dict['Rod Dia'] = re.search(r'Dia, in:(.*?)Mfg', text).group(1).strip()
if re.search(r'Mfg:(.*)', text).group(1):
    output_dict['Driver Mfg'] = re.search(r'Mfg:(.*)', text).group(1)
else:
    output_dict['Driver Mfg'] = "none"

output_dict['Max RL Tot'] = re.search(r'Max RL Tot, lbf:(.*?)Max RL Tens', text).group(1).strip()
output_dict['Max RL Tens'] = re.search(r'Max RL Tens, lbf:(.*?)Max RL Comp', text).group(1).strip()
output_dict['Max RL Comp'] = re.search(r'Max RL Comp, lbf:(.*?)Model', text).group(1).strip()
if re.search(r'Model:(.*)', text).group(1):
    output_dict['Driver Model'] = re.search(r'Model:(.*)', text).group(1)
else:
    output_dict['Driver Model'] = 'none'

output_dict['Rated RPM'] = re.search(fr'Rated RPM:(.*?)Rated {power}', text).group(1).strip()
output_dict[f'Rated {power}'] = re.search(fr'Rated {power}:(.*?)Rated PS', text).group(1).strip()
output_dict['Rated PS'] = re.search(fr'Rated PS FPM:(.*?){power}', text).group(1).strip()
output_dict[f'{power}'] = re.search(fr'\d {power}:(.*)', text).group(1).strip()

output_dict['Calc RPM'] = re.search(fr'Calc RPM:(.*?){power}:', text).group(1).strip()
output_dict[f'Calc {power}'] = re.search(fr'{power}:(.*?)Calc PS', text).group(1).strip()
output_dict['Calc PS'] = re.search(r'Calc PS FPM:(.*?)Avail', text).group(1).strip()
output_dict['Avail HP'] = re.search(r'Avail:(.*)', text).group(1).strip()

output_dict['Services'] = re.search(r'Services(.*)', text).group(1).strip()
output_dict['Gas Model'] = re.search(r'Gas Model(.*)', text).group(1).strip()
output_dict['Stage Data'] = re.search(r'Stage Data:(.*)', text).group(1).strip()
output_dict[f'Target Flow, {flow}'] = re.search(fr'Target Flow, {flow}(.*)', text).group(1).strip()
output_dict[f'Flow Calc, {flow}'] = re.search(fr'Flow Calc, {flow}(.*)', text).group(1).strip()
output_dict[f'{power} per Stage'] = re.search(fr'{power} per Stage(.*)', text).group(1).strip()
output_dict['Specific Gravity'] = re.search(r'Specific Gravity(.*)', text).group(1).strip()
output_dict['Ratio of Sp Ht (N)'] = re.search(r'Ratio of Sp Ht \(N\)(.*)', text).group(1).strip()
output_dict['Comp Suct (Zs)'] = re.search(r'Comp Suct \(Zs\)(.*)', text).group(1).strip()
output_dict['Comp Disch (Zd)'] = re.search(r'Comp Disch \(Zd\)(.*)', text).group(1).strip()
output_dict[f'Pres Suct Line, {pressure}'] = re.search(fr'Pres Suct Line, {pressure}(.*)', text).group(1).strip()
output_dict[f'Pres Suct Flg, {pressure}'] = re.search(fr'Pres Suct Flg, {pressure}(.*)', text).group(1).strip()
output_dict[f'Pres Disch Flg, {pressure}'] = re.search(fr'Pres Disch Flg, {pressure}(.*)', text).group(1).strip()
output_dict[f'Pres Disch Line, {pressure}'] = re.search(fr'Pres Disch Line, {pressure}(.*)', text).group(1).strip()
output_dict['Pres Ratio F/F'] = re.search(r'Pres Ratio F/F(.*)', text).group(1).strip()
output_dict[f'Temp Suct, {temperature}'] = re.search(fr'Temp Suct, {temperature}(.*)', text).group(1).strip()
output_dict[f'Temp Clr Disch, {temperature}'] = re.search(fr'Temp Clr Disch, {temperature}(.*)', text).group(1).strip()
output_dict['Cylinder Data'] = re.search(r'Cylinder Data:(.*)', text).group(1).strip()
output_dict['Cyl Model'] = re.search(r'Cyl Model(.*)', text).group(1).strip()
output_dict['Cyl Bore, in'] = re.search(r'Cyl Bore, in(.*)', text).group(1).strip()
output_dict[f'Cyl RDP (API), {pressure}'] = re.search(fr'Cyl RDP \(API\), {pressure}(.*)', text).group(1).strip()
output_dict[f'Cyl MAWP, {pressure}'] = re.search(fr'Cyl MAWP, {pressure}(.*)', text).group(1).strip()
output_dict['Cyl Action'] = re.search(r'Cyl Action(.*)', text).group(1).strip()
output_dict['Cyl Disp, CFM'] = re.search(r'Cyl Disp, CFM(.*)', text).group(1).strip()
output_dict[f'Pres Suct Intl, {pressure}'] = re.search(fr'Pres Suct Intl, {pressure}(.*)', text).group(1).strip()
output_dict['Temp Suct Intl, F'] = re.search(r'Temp Suct Intl, F(.*)', text).group(1).strip()
output_dict[f'Pres Disch Intl, {pressure}'] = re.search(fr'Pres Disch Intl, {pressure}(.*)', text).group(1).strip()
output_dict[f'Temp Disch Intl, {temperature}'] = re.search(fr'Temp Disch Intl, {temperature}(.*)', text).group(1).strip()
output_dict['HE Suct Gas Vel, FPM'] = re.search(r'HE Suct Gas Vel, FPM(.*)', text).group(1).strip()
output_dict['HE Disch Gas Vel, FPM'] = re.search(r'HE Disch Gas Vel, FPM(.*)', text).group(1).strip()
output_dict['HE Spcrs Used/Max'] = re.search(r'HE Spcrs Used/Max(.*)', text).group(1).strip()
output_dict['HE Vol Pkt Avail'] = re.search(r'HE Vol Pkt Avail(.*)', text).group(1).strip()
output_dict['Vol Pkt Used'] = re.search(r'Vol Pkt Used(.*)', text).group(1).strip()
output_dict['HE Min Clr, %'] = re.search(r'HE Min Clr, %(.*)', text).group(1).strip()
output_dict['HE Total Clr, %'] = re.search(r'HE Total Clr, %(.*)', text).group(1).strip()
output_dict['CE Suct Gas Vel, FPM'] = re.search(r'CE Suct Gas Vel, FPM(.*)', text).group(1).strip()
output_dict['CE Disch Gas Vel, FPM'] = re.search(r'CE Disch Gas Vel, FPM(.*)', text).group(1).strip()
output_dict['CE Spcrs Used/Max'] = re.search(r'CE Spcrs Used/Max(.*)', text).group(1).strip()
output_dict['CE Min Clr, %'] = re.search(r'CE Min Clr, %(.*)', text).group(1).strip()
output_dict['CE Total Clr, %'] = re.search(r'CE Total Clr, %(.*)', text).group(1).strip()
output_dict['Suct Vol Eff HE/CE, %'] = re.search(r'Suct Vol Eff HE/CE, %(.*)', text).group(1).strip()
output_dict['Disch Event HE/CE, ms'] = re.search(r'Disch Event HE/CE, ms(.*)', text).group(1).strip()
output_dict['Suct Pseudo-Q HE/CE'] = re.search(r'Suct Pseudo-Q HE/CE(.*)', text).group(1).strip()
output_dict['Gas Rod Ld Comp, %'] = re.search(r'Gas Rod Ld Comp, %(.*)', text).group(1).strip()
output_dict['Gas Rod Ld Tens, %'] = re.search(r'Gas Rod Ld Tens, %(.*)', text).group(1).strip()
output_dict['Gas Rod Ld Total, %'] = re.search(r'Gas Rod Ld Total, %(.*)', text).group(1).strip()
output_dict['Xhd Pin Deg/%Rvrsl lbf'] = re.search(r'Xhd Pin Deg/%Rvrsl lbf(.*)', text).group(1).strip()
# output_dict['Flow Calc, MMSCFD'] = re.search(r'Flow Calc, MMSCFD(.*)', text).group(1).strip()
output_dict[f'Cyl {power}'] = re.search(fr'Cyl {power}(.*)', text).group(1).strip()


#######################################
# Now that our text string is broken up into a useful dict, we can get to key
#  information like how many services this machine has, how many cylinders
#  the compressor has, how many are associated with each service, each stage,
#  and which cylinders those are (bore, nominal, mawp).  In this section we
#  mine out this information and add it ot the stages dict.
services_lst = [service for service in output_dict['Services'].split() if service.isdigit()]
# print(services_lst)
num_services = services_lst[-1]
# print(num_services)
#############################
#############################
# This is the problem.  The split on 1 drops the last service b/c its a single stage
# cyls_list = [f'1{stage}'.split() for stage in output_dict['Stage Data'].split('1') if stage]
cyls_list = [f'1{stage}'.split() for stage in output_dict['Stage Data'].split('1')][1:]
for index in range(len(cyls_list)):
    cyls_list[index] = [cyl for cyl in cyls_list[index] if cyl != '(SG)']
print(cyls_list)
stages_list = {}
for index in range(len(cyls_list)):
    stages_list[index] = [stage for stage in cyls_list[index] if stage != '---']
print(stages_list)

stg_data = output_dict[f'Stage Data'].split()
stg_data.remove('(SG)')
cylinders = []
for index in range(len(output_dict['Cyl Model'].split())):
    cylinders.append([
    output_dict['Cyl Model'].split()[index],
    output_dict['Cyl Bore, in'].split()[index],
    output_dict[f'Cyl RDP (API), {pressure}'].split()[index],
    output_dict[f'Cyl MAWP, {pressure}'].split()[index],
    'Throw '+ output_dict[f'Cylinder Data'].split('Throw')[index + 1],
    ])
    # if stg_data[index] != '---':
    #     cylinders[index].append('Stage ' + stg_data[index])
    # else:
    #     temp_index = index - 1
    #     if stg_data[temp_index] != '---':
    #         cylinders[index].append('Stage ' + stg_data[temp_index])

    print(cylinders[index])
# print(stg_data)

stages = {}
stages[f'Total Services'] = num_services
for service in range(int(num_services)):
    stages[f'Service {int(service) + 1} Total Stages'] = int(stages_list[service][-1])
    stages[f'Service {int(service) + 1} Total Cylinders'] = len(cyls_list[service])
    current_stage = 1
    for throw in cyls_list[service]:
        if throw.isdigit():
            stages[f'Service {int(service) + 1} Stage {throw}'] = 1
            current_stage = throw
        else:
            stages[f'Service {int(service) + 1} Stage {current_stage}'] += 1


column_location = 0
for service in range(int(num_services)):
    stages[f'Service {service + 1} Column Start'] = column_location
    for stage in range(int(stages[f'Service {service + 1} Total Stages'])):
        stages[f'Service {service + 1} Stage {int(stage + 1)} Cylinder'] = cylinders[column_location]
        stages[f'Service {service + 1} Stage {int(stage + 1)} Column Start'] = column_location
        column_location += stages[f'Service {service + 1} Stage {int(stage + 1)}']



# In this section we define a few functions that either manipulate our data
#  into the format the runM requires (pressure in absolute, temp in R) or
#  populate our final output.


# pressure_corrector takes the start column of a service and the total stages
#  that service utilizes as arguments, then checks if pressure is in gauge or
#  absolute.  If in gauge, pressure is converted to absolute and the function
#  returns a list ps_pd with suction ps as the 1st item and pd as the 2nd.
# ***This function does not consistently work***
#   there is something going on with the way the runM converts abs back to g
#   when building the run file that is inconsistent.  It doesn't just subtract 14.7
#   to the provided pressure value.  value is typically somewhere between 13.5 and
#   14.7.  Need to clarify.

def pressure_corrector(col_start, tot_cyl):
    ps_pd = ["", ""]
    if g_or_abs == 'Gauge':
        ps_pd[0] = float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]) + 13.3
        pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
        print(pd_lst)
        ps_pd[1] = float(pd_lst[-1]) + 13.3
        return ps_pd
    else:
        ps_pd[0] = float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start])
        ps_pd[1] = float(output_dict[f'Pres Disch Line, {pressure}'].split() [col_start+tot_cyl - 1])


# The stage_assigner takes the service as an argument and populates an output
#  string with each cylinder of that stage in the runM format.
#  Need to talk with the guys about what loss_factor is and how to make this
#  variable from the report.  currently hard coded as 0.00704
def stage_assigner(service):
    output = ""
    for stage in range(int(stages[f'Service {service} Total Stages'])):
        output = output + (f"""<Stage>
                	<number>{stage + 1}</number>
                	<Cylinder>
                		<total>{stages[f'Service {service} Stage {stage + 1}']}</total>
                		<product_family>{product_family[output_dict['Frame'][-5:-2]]}</product_family>
                		<bore_size>{stages[f'Service {service} Stage {stage + 1} Cylinder'][1]}</bore_size>
                		<loss_factor>0.00704</loss_factor>
                	</Cylinder>
                </Stage>
            """)
    return output

# temp_converter converts any temp to R.  Seems to consistently work.
def temp_converter(suct_temp):
    if temperature =="F":
        suct_temp_converted = round(suct_temp + 459.67, 2)
    elif temperature == "C":
        suct_temp_converted = round(suct_temp * 9/5 + 491.67, 2)
    elif temperature == "K":
        suct_temp_converted = round(suct_temp * 1.8, 2)
    else:
        suct_temp_converted = suct_temp
    return suct_temp_converted

# gas_type looks at the sg and takes a guess at the gas_type based on that.
#  Future improvement might be to search the text string for 'Sour Gas' and
#  select sour gas in that scenario
def gas_type(sg):
    sgf = float(sg)
    if sgf >= 0.6000 and sgf <= 1.1000:
        return 'Field Gas'
    elif sgf > 1.1000 and sgf <=2.0000:
        return 'Heavy Hydrocarbons'
    elif sgf >= 0.5500 and sgf <0.6000:
        return 'Pipeline Quality'

# service_assigner loops through each service of the machine and populates an
#  output string with the required information in runM format.  it calls the
#  stage_assigner function for each service to populate the stages portion and
#  calls some of the other functions to populate the particular value in question.
def service_assigner():
    output = ""
    for service in range(1, int(stages[f'Total Services']) + 1):
        output = output + f"""
        <Service>
            <target_flow>{output_dict[f'Flow Calc, {flow}'].split()[stages[f'Service {service} Column Start']]}</target_flow>
            <target_bhp>{output_dict[f'Calc {power}']}</target_bhp>
            <suction_pressure>{pressure_corrector(stages[f'Service {service} Column Start'], stages[f'Service {service} Total Cylinders'])[0]}</suction_pressure>
            <discharge_pressure>{pressure_corrector(stages[f'Service {service} Column Start'], stages[f'Service {service} Total Cylinders'])[1]}</discharge_pressure>
            <suction_temp>{temp_converter(float(output_dict[f'Temp Suct, {temperature}'].split()[stages[f'Service {service} Column Start']]))}</suction_temp>
            <gas_type>{gas_type(output_dict['Specific Gravity'].split()[stages[f'Service {service} Column Start']])}</gas_type>
            <specific_gravity>{output_dict['Specific Gravity'].split()[stages[f'Service {service} Column Start']]}</specific_gravity>
            <k>{output_dict['Ratio of Sp Ht (N)'].split()[stages[f'Service {service} Column Start']]}</k>
            <Zs>{output_dict['Comp Suct (Zs)'].split()[stages[f'Service {service} Column Start']]}</Zs>
            <Zd>{output_dict['Comp Disch (Zd)'].split()[stages[f'Service {service} Column Start']]}</Zd>
            <total_stages>{stages[f'Service {service} Total Stages']}</total_stages>
            <Stages>
                {stage_assigner(service)}
            </Stages>
    </Service>"""
    return output


# Finally the program creates an output_txt string which creates the top portion
#  of the runM format and then calls the service assigner to fill in service
#  and stage specific information.

output_txt = fr"""<?xml version="1.0" encoding=UTF-8 ?>
<Ariel_Mobile>
    <Version>1.0</Version>
    <Units>
        <flow>{flow}</flow>
        <power>{power_for_output}</power>
        <temperature>{temperature}</temperature>
        <pressure>{pressure_for_output}</pressure>
        <unitType>{eng_met}</unitType>
        <pressType>{g_or_abs}</pressType>
    </Units>
    <User>
		<name>NA</name>
		<email>NA</email>
		<telephone>NA</telephone>
		<project>{output_dict['Project']}</project>
		<location>NA</location>
		<country>NA</country>
		<elevation>{output_dict["Elevation"]}</elevation>
	</User>
	<Compressor>
		<RPM>{output_dict['Calc RPM']}</RPM>
		<product_family>{product_family[output_dict['Frame'][-5:-2]]}</product_family>
		<model>{output_dict['Frame'][-5:-2]}</model>
		<num_throws>{output_dict['Frame'][-1]}</num_throws>
	</Compressor>
    <Services>
	{service_assigner()}
    </Services>
</Ariel_Mobile>
"""

# print(output_txt)

# Current program limitaions include inability to select cylinders or cylinder
#  action / clearance.  This means that the program will autoselect the cylinders
#  attempting to match whatever volumen is given to target flow.  Currently,
#  I have the target flow value set to the calc flow from the input .pdf so
#  the program at least attempts to size the cylinders to match the actual
#  throughput on the .pdf.  This if flawed b/c if the input .pdf has the first
#  stg single acted for example, the program will select smaller cylinders
#  goal seeking the volume from that configuration.
#  Another limitation is that we can't set the ambient, or the cooler discharge
#  temps, or any of the losses.
#  Also need to figure out what is going on with the pressure conversion b/c
#  pressures don't come out just right consistently (although usually within a
#  pound or two)


with open(save_filepath, mode='w', encoding='utf-8') as file:
    file.writelines(output_txt)

os.startfile(save_filepath)
