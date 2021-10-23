import pdfplumber
import re
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import os
import traceback


# open_filepath = r'C:\Users\Jay\Desktop\Python\Auto Run File Builder\Tests\Test_15_F_9.PDF'
# save_filepath = r'C:\Users\Jay\Desktop\Python\Auto Run File Builder\Tests\Test_15_F_9.runM'

open_filepath = None
save_filepath = None


def run_calcs():
    #I've got the bulk of the code in this try / except block so that I can pop up an
    #  error window for the user to see if the program doesn't work before the runM is
    #  created and opened.
    try:
        # Opens the run report using the pdfplumber module, extracts the text and
        #  saves it to one big string called 'text'
        with pdfplumber.open(open_filepath) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()

        # print(text)


        # here we search the text string for particular attributes using regular
        #  expressions.  We need many of these attributes to create the output_dict
        # later.  For example, the variable 'flow' here will return the flow
        #  units, like MMSCFD, which we need to know which units are being utilized
        #  as well as to effectively search the rest of the text string for other
        #  values ie Calc Flow {flow}
        flow = re.search(r'Target Flow, (.*?) \d', text).group(1).strip()
        power = re.search(f'Rated RPM: \d+ Rated (\D+): \d+', text).group(1).strip()
        power_for_output = power[-2:].lower()
        temperature = re.search(r'Ambient,(.*?):', text).group(1).strip()
        pressure = re.search(r'Pres Suct Line, (.*?) \d', text).group(1).strip()
        if re.search(r'Max RL Tot, (.*?):', text).group(1).strip() == 'lbf':
            eng_met = 'English'
            length = 'ft'
            small_length = 'in'
            force = 'lbf'
            piston_speed = 'FPM'
            displacement = 'CFM'
        else:
            eng_met = 'Metric'
            length = 'm'
            small_length = 'mm'
            force = 'kN'
            piston_speed = 'm/s'
            displacement = 'm3/h'

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
        if pressure == 'atm':
            output_dict[f'Barmtr,{pressure}'] = re.search(fr'Barmtr,{pressure}:(.*?)Ambient', text).group(1).strip()
        else:
            output_dict[f'Barmtr,{pressure[:1]+"a"}'] = re.search(fr'Barmtr,{pressure[:-1]+"a"}:(.*?)Ambient', text).group(1).strip()
        output_dict[f'Ambient,{temperature}'] = re.search(fr'Ambient,{temperature}:(.*?)Type', text).group(1).strip()
        output_dict['Driver Type'] = re.search(r'Type:(.*)', text).group(1).strip()

        output_dict['Frame'] = re.search(r'Frame:(.*?)Stroke', text).group(1).strip()
        output_dict['Stroke'] = re.search(fr'Stroke, {small_length}:(.*?)Rod', text).group(1).strip()
        output_dict['Rod Dia'] = re.search(fr'Dia, {small_length}:(.*?)Mfg', text).group(1).strip()

        # driver MFG and Model require a bit of logic since they are the only two values that
        #  can be left blank.  When blank, we manually input the 'none' string as the value
        if re.search(r'Mfg:(.*)', text).group(1):
            output_dict['Driver Mfg'] = re.search(r'Mfg:(.*)', text).group(1)
        else:
            output_dict['Driver Mfg'] = "none"

        output_dict['Max RL Tot'] = re.search(fr'Max RL Tot, {force}:(.*?)Max RL Tens', text).group(1).strip()
        output_dict['Max RL Tens'] = re.search(fr'Max RL Tens, {force}:(.*?)Max RL Comp', text).group(1).strip()
        output_dict['Max RL Comp'] = re.search(fr'Max RL Comp, {force}:(.*?)Model', text).group(1).strip()
        if re.search(r'Model:(.*)', text).group(1):
            output_dict['Driver Model'] = re.search(r'Model:(.*)', text).group(1)
        else:
            output_dict['Driver Model'] = 'none'

        output_dict['Rated RPM'] = re.search(fr'Rated RPM:(.*?)Rated {power}', text).group(1).strip()
        output_dict[f'Rated {power}'] = re.search(fr'Rated {power}:(.*?)Rated PS', text).group(1).strip()
        output_dict['Rated PS'] = re.search(fr'Rated PS {piston_speed}:(.*?){power}', text).group(1).strip()
        output_dict[f'{power}'] = re.search(fr'\d {power}:(.*)', text).group(1).strip()


        output_dict['Calc RPM'] = re.search(fr'Calc RPM:(.*?){power}:', text).group(1).strip()
        output_dict[f'Calc {power}'] = re.search(fr'{power}:(.*?)Calc PS', text).group(1).strip()
        output_dict['Calc PS'] = re.search(fr'Calc PS {piston_speed}:(.*?)Avail', text).group(1).strip()
        output_dict['Avail HP'] = re.search(r'Avail:(.*)', text).group(1).strip()

        #Here we put in some logic to handle the situation when an amound of power is included.
        # if power and a derate is specificd it will be in the format "100 (8%)".  when we split
        # that string, since there is a space between power and the derate parentheses this results in
        # a list of length 2. In that case, we split the power number from the derate percent
        # and then divide the power number by 1 - Derate % to get total power at that RPM.
        # Then we select the derate and aux values from the power and available sections
        # (both in index 1 of their respective split lists).  If the power number and/or Avail HP
        # is the only item in that spot, the list resulting from len(output_dict[f'{power}'].split())
        # or Avail HP is 1 and derate and aux are set to zero.
        if len(output_dict[f'{power}'].split()) > 1:
            driver_rated_BHP = int(float(output_dict[f'{power}'].split()[0]) / (1 - float(output_dict[f'{power}'].split()[1].replace("(", "").replace(")", "").replace("%", ""))/100))
            driver_derate = output_dict[f'{power}'].split()[1].replace("(", "").replace(")", "")
            # driver_aux = output_dict['Avail HP'].split()[1].replace("(", "").replace(")", "")
        else:
            driver_rated_BHP = output_dict[f'{power}']
            driver_derate = "0"
            # driver_aux = "0"

        if len(output_dict[f'Avail HP'].split()) > 1:
            driver_aux = int(output_dict['Avail HP'].split()[1].replace("(", "").replace(")", ""))
        else:
            driver_aux = 0

        if power_for_output == 'kw':
            driver_rated_BHP = round(int(driver_rated_BHP) * 1.34102, 0)
            driver_aux = round(int(driver_aux) * 1.34102, 0)

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
        # the Cyl Model and Pkt lines are a little different as they can wrap to the next line in certain circumstances
        # the original logic is commented out below followed by logic that collects text wrapped to the next line.
        # Need further clarification about how wraps work in program output (same column next line or is
        # everything pushed over one column and the far right info wrapped to the next line)
        #######################
        #######################
        # this should probably be reworked into a function
        # output_dict['Cyl Model'] = re.search(r'Cyl Model(.*)', text).group(1).strip()
        # output_dict['Cyl Model'] = re.search(r'Cyl Model(.*)Cyl Bore', text, re.DOTALL).group(1).replace("\n", "").strip()
        output_dict['Cyl Model'] = re.search(r'Cyl Model(.*)Cyl Bore', text, re.DOTALL).group(1).strip()
        if "\n" in output_dict['Cyl Model']:
            cyl_model_line_split = [item.split() for item in output_dict['Cyl Model'].split("\n")]
            counter = 0
            for indx in range(len(cyl_model_line_split[0])):
                if cyl_model_line_split[0][indx].endswith('-'):
                    cyl_model_line_split[0][indx] = cyl_model_line_split[0][indx] + cyl_model_line_split[1][counter]
                    counter +=1
            output_dict['Cyl Model'] = " ".join(cyl_model_line_split[0])

        output_dict[f'Cyl Bore, {small_length}'] = re.search(fr'Cyl Bore, {small_length}(.*)', text).group(1).strip()
        if pressure == "atm":
            output_dict[f'Cyl RDP (API), {pressure}'] = re.search(fr'Cyl RDP \(API\), {pressure}(.*)', text).group(1).strip()
            output_dict[f'Cyl MAWP, {pressure}'] = re.search(fr'Cyl MAWP, {pressure}(.*)', text).group(1).strip()
        else:
            output_dict[f'Cyl RDP (API), {pressure[:-1]+"g"}'] = re.search(fr'Cyl RDP \(API\), {pressure[:-1]+"g"}(.*)', text).group(1).strip()
            output_dict[f'Cyl MAWP, {pressure[:-1]+"g"}'] = re.search(fr'Cyl MAWP, {pressure[:-1]+"g"}(.*)', text).group(1).strip()
        output_dict['Cyl Action'] = re.search(r'Cyl Action(.*)', text).group(1).strip()
        output_dict[f'Cyl Disp, {displacement}'] = re.search(fr'Cyl Disp, {displacement}(.*)', text).group(1).strip()
        output_dict[f'Pres Suct Intl, {pressure}'] = re.search(fr'Pres Suct Intl, {pressure}(.*)', text).group(1).strip()
        output_dict[f'Temp Suct Intl, {temperature}'] = re.search(fr'Temp Suct Intl, {temperature}(.*)', text).group(1).strip()
        output_dict[f'Pres Disch Intl, {pressure}'] = re.search(fr'Pres Disch Intl, {pressure}(.*)', text).group(1).strip()
        output_dict[f'Temp Disch Intl, {temperature}'] = re.search(fr'Temp Disch Intl, {temperature}(.*)', text).group(1).strip()
        output_dict[f'HE Suct Gas Vel, {piston_speed}'] = re.search(fr'HE Suct Gas Vel, {piston_speed}(.*)', text).group(1).strip()
        output_dict[f'HE Disch Gas Vel, {piston_speed}'] = re.search(fr'HE Disch Gas Vel, {piston_speed}(.*)', text).group(1).strip()
        output_dict['HE Spcrs Used/Max'] = re.search(r'HE Spcrs Used/Max(.*)', text).group(1).strip()
        output_dict['HE Vol Pkt Avail'] = re.search(r'HE Vol Pkt Avail(.*)', text).group(1).strip()
        # Same as mentioned above, using DOTALL flag to grab everything, including new lines, between
        # the search texts.
        # output_dict['Vol Pkt Used'] = re.search(r'Vol Pkt Used(.*)', text).group(1).strip()
        output_dict['Vol Pkt Used'] = re.search(r'Vol Pkt Used(.*)HE Min Clr', text, re.DOTALL).group(1).replace("\n", " ").strip()
        # print(output_dict['Vol Pkt Used'])
        output_dict['HE Min Clr, %'] = re.search(r'HE Min Clr, %(.*)', text).group(1).strip()
        output_dict['HE Total Clr, %'] = re.search(r'HE Total Clr, %(.*)', text).group(1).strip()
        output_dict[f'CE Suct Gas Vel, {piston_speed}'] = re.search(fr'CE Suct Gas Vel, {piston_speed}(.*)', text).group(1).strip()
        output_dict[f'CE Disch Gas Vel, {piston_speed}'] = re.search(fr'CE Disch Gas Vel, {piston_speed}(.*)', text).group(1).strip()
        output_dict['CE Spcrs Used/Max'] = re.search(r'CE Spcrs Used/Max(.*)', text).group(1).strip()
        output_dict['CE Min Clr, %'] = re.search(r'CE Min Clr, %(.*)', text).group(1).strip()
        output_dict['CE Total Clr, %'] = re.search(r'CE Total Clr, %(.*)', text).group(1).strip()
        output_dict['Suct Vol Eff HE/CE, %'] = re.search(r'Suct Vol Eff HE/CE, %(.*)', text).group(1).strip()
        output_dict['Disch Event HE/CE, ms'] = re.search(r'Disch Event HE/CE, ms(.*)', text).group(1).strip()
        output_dict['Suct Pseudo-Q HE/CE'] = re.search(r'Suct Pseudo-Q HE/CE(.*)', text).group(1).strip()
        output_dict['Gas Rod Ld Comp, %'] = re.search(r'Gas Rod Ld Comp, %(.*)', text).group(1).strip()
        output_dict['Gas Rod Ld Tens, %'] = re.search(r'Gas Rod Ld Tens, %(.*)', text).group(1).strip()
        output_dict['Gas Rod Ld Total, %'] = re.search(r'Gas Rod Ld Total, %(.*)', text).group(1).strip()
        output_dict[f'Xhd Pin Deg/%Rvrsl {force}'] = re.search(fr'Xhd Pin Deg/%Rvrsl {force}(.*)', text).group(1).strip()
        # the 'Flow Calc, {power}' row heading is used twice on the perf report, once to denote
        # stage flow calc and again on a per cylinder basis.  Rather than distinguish between the two,
        # the below is just commented out.  Can address later if required.
        # output_dict['Flow Calc, {power}'] = re.search(r'Flow Calc, MMSCFD(.*)', text).group(1).strip()
        output_dict[f'Cyl {power}'] = re.search(fr'Cyl {power}(.*)', text).group(1).strip()


        #######################################
        # Now that our text string is broken up into a useful dict, we can get to key
        #  information like how many services this machine has, how many cylinders
        #  the compressor has, how many are associated with each service, each stage,
        #  and which cylinders those are (bore, nominal, mawp).  In this section we
        #  mine out this information and add it ot the stages dict.

        # services_lst = [service for service in output_dict['Services'].split() if service.isdigit()]
        services_lst = [stage for stage in output_dict['Stage Data'].split() if stage == '1']
        #Aparently on some run files, the services line wouldn't include any numbers, at least
        # in some single stg applications.  Rather than showing "Services Service 1" the run report would
        # read "Services Gathering".  Not sure how this worked for multiservice applications but
        # to handle the example I have the logic below checks to see if the services list was
        # populated with anything (using the .isdigit()) and if not adds a service to the list
        # b/c every application will have at least one service.
        # Need more information of how multi service applications would be handled in this vintage of
        # the performance software.
        # if len(services_lst) < 1:
        #     services_lst = [1]
        # print(services_lst)
        # print(len(services_lst))

        # num_services = services_lst[-1]
        num_services = len(services_lst)
        # print(num_services)
        #############################
        #############################
        # This was a problem.  The split on 1 drops the last service b/c its a single stage
        # Since there will always be at least 1 stage in each service, hard code "1" into the string.
        # cyls_list is a list comprehension.  We take the stage data line from the run report
        # and split it at the number 1.  The split is not inclusive so we end up with something
        # like ['', '   (SG) --- 2 3 4 ', '   (SG)'] which would represent a 2 service
        # machine with 4 stages on the 1st service and 1 stage on the 2nd service.
        # splitting on the leading character puts in an empty list entry to represent
        # what was ahead of the leading character so we slice the list from the first character
        # on using [1:].  Then in the list comprehension we add '1' to each list item to replace
        # the '1' that we split on (not inclusive) and then split that entry.  So the resultant list
        # is a list of lists, in this case [['1', '(SG)', '---', '2', '3', '4'], ['1', '(SG)']].
        # then we iterate through each list in the cyls_list list and using another list comprehension
        # keep each item that is not '(SG)'
        cyls_list = [f'1{stage}'.split() for stage in output_dict['Stage Data'].split('1')][1:]
        # print(cyls_list)

        for index in range(len(cyls_list)):
            cyls_list[index] = [cyl for cyl in cyls_list[index] if cyl != '(SG)']
        # print('cyls_list: ', cyls_list)
        stages_list = {}
        for index in range(len(cyls_list)):
            stages_list[index] = [stage for stage in cyls_list[index] if stage != '---']
        # print('stages_list: ', stages_list)

        #The stg_data and stg_data_checker are used to keep track of multi-cylinder stages
        # the stg_data list is basically the same list at the cyls_list except it is not
        # broken up into sub-lists per service.  Could possibly change this to just join
        # the cyls_list but the code below also works.  Then the stg_data_checker function
        # recursively checks if the column in question has a '---' entry for the stage data,
        # if it does, we move back one column and run the function again, if the stage entry
        # is a number we know the cylinder in that column is the first or only one in that stage
        # and we can gather the relavent stage data from that column.

        stg_data = [stg for stg in output_dict[f'Stage Data'].split() if stg != '(SG)']
        # print('stg_data: ', stg_data)

        def stg_data_checker(pos):
            if stg_data[pos] == '---':
                pos = pos - 1
                return stg_data_checker(pos)
            else:
                return 'Stage ' + stg_data[pos]

        #This logic uses re.split to search the output_dict volume pocket line
        # on multiple delimiters to split the pocket line into useable information.
        # since the pocket units can be of several different types but the type is
        # not called out in the row heading, we search on all possible options.
        # The re.split search is inclusive of the delimiter but it splits the
        # delimiter off from the item so we search the list for every delimiter
        # and when we find one, we add the delimiter to the list item immediately
        # before that list position.  Then we use a list comprehension to drop all
        # the delimiter items that are not connected to a value.
        if eng_met == 'English':
            pkt_used = re.split(r"(%|turns|in|Pos, in|No Pkt)", output_dict['Vol Pkt Used'])
            for index in range(len(pkt_used)):
                for delim in ["%", "turns", "in", "Pos, in"]:
                    if pkt_used[index] == delim:
                        pkt_used[index-1] += delim
            pkt_used = [item.strip() for item in pkt_used if item.strip() and item not in ["%", "turns", "in", "Pos, in"]]
        else:
            pkt_used = re.split(r"(%|turns|mm|Pos, cm|No Pkt)", output_dict['Vol Pkt Used'])
            for index in range(len(pkt_used)):
                for delim in ["%", "turns", "mm", "Pos, cm"]:
                    if pkt_used[index] == delim:
                        pkt_used[index-1] += delim
            pkt_used = [item.strip() for item in pkt_used if item.strip() and item not in ["%", "turns", "mm", "Pos, cm"]]

        #Here we loop through the cylinder model line and create a list with the cylinder specific
        # information for every throw of the machine.  We have to use logic to handle the
        # bore and MAWP/RDP since these can vary dependent on the units used.
        # The cylinders list ends up being a list of lists.
        cylinders = []
        # 0-model, 1-bore, 2-rdp, 3-mawp, 4-throw, 5-stage, 6-action, 7-he spcrs, 8-ce spcrs, 9-pkt avail, 10-pkt used
        for index in range(len(output_dict['Cyl Model'].split())):
            cylinders.append([
            output_dict['Cyl Model'].split()[index],
            ])
            if eng_met == "English":
                cylinders[index].append(output_dict[f'Cyl Bore, {small_length}'].split()[index])
            else:
                cylinders[index].append(f"{float(output_dict[f'Cyl Bore, {small_length}'].split()[index]) * 0.0393701:.3f}")
            if pressure == 'atm':
                cylinders[index].append(output_dict[f'Cyl RDP (API), {pressure}'].split()[index])
                cylinders[index].append(output_dict[f'Cyl MAWP, {pressure}'].split()[index])
            else:
                cylinders[index].append(output_dict[f'Cyl RDP (API), {pressure[:-1]+"g"}'].split()[index])
                cylinders[index].append(output_dict[f'Cyl MAWP, {pressure[:-1]+"g"}'].split()[index])
            cylinders[index].append(output_dict[f'Cylinder Data'].split('Throw')[index + 1])

            cylinders[index].append(stg_data_checker(index))
            cylinders[index].append(output_dict['Cyl Action'].split()[index])
            cylinders[index].append(output_dict['HE Spcrs Used/Max'].split()[index][0])
            cylinders[index].append(output_dict['CE Spcrs Used/Max'].split()[index][0])
            cylinders[index].append(output_dict['HE Vol Pkt Avail'].replace('No Pkt', 'No_Pkt').split()[index].replace('No_Pkt', 'No Pkt'))
            cylinders[index].append(pkt_used[index])

        # for cyl in cylinders:
        #     print(cyl)

        #Here we create the stages dict which will hold all of the relavent information
        # about each stage.  This will allow us to programatically call this information
        # for any combination of stage and service when we are creating the output text.
        # First, we capture the total number of services, then for each service we
        # loop through and capture the number to stages per service and the number of
        # cylinders per service.  Next we look at the index of the cyls_list list
        # that corresponds to our current service.  Each entry of the cyls_list
        # will be something like ['1', '---', '---', '---'], which would represent
        # a 4 throw single stage.  So we walk through the entry and if the list item
        # is a number, we set the 'Service x Stage y' dict key to the value 1.  Then
        # we set the current_stage to the digit which is the 'throw' in our loop.
        # If the next list item is a non-digit, it '---', we add 1 to the
        # 'Service x Stage y' dict key's value but do not change the current_stage.
        # This allows us to use the 'Service x Stage {current_stage}' key to set the
        # value which is handy since using 'Service x Stage {throw}' like we did
        # initially wouldn't work since throw counts up each step regardless of
        # weather it's a digit or not.  We add 1 to the Service x Stage y key
        # for every step until we reach another digit and when another digit is
        # reached we set a new Service x Stage y key equal to 1 and repeat the process.
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


        #Here we create a counter variable column_location and step through the columns
        # of the performance report.  For each service we set which column it starts in
        # then for each stage in that service (already added to the stages dict above)
        # we assing relavent information to the stages dict.  We create a the temp_loc
        # counter and set it equal to column_location current value.  Then we add the
        # total number of cylinders in that stage to column_location.  Next we create
        # a number of local variables to keep track of throws, action spacers etc and
        # assign them the respective values in the temp_loc's column before incrementing
        # temp_loc by 1.  We then recursively add column information to these local variables
        # until the temp_loc equals the new column_location at which point we assign each
        # local variable to a key in the stages dict.
        column_location = 0
        for service in range(int(num_services)):
            stages[f'Service {service + 1} Column Start'] = column_location
            for stage in range(int(stages[f'Service {service + 1} Total Stages'])):
                stages[f'Service {service + 1} Stage {int(stage + 1)} Cylinder'] = [cylinders[column_location]]
                stages[f'Service {service + 1} Stage {int(stage + 1)} Column Start'] = column_location
                stages[f'Service {service +1} Stage {int(stage + 1)} Suction Temp'] = output_dict[f'Temp Suct, {temperature}'].split()[column_location]
                stages[f'Service {service +1} Stage {int(stage + 1)} Cooler Temp'] = output_dict[f'Temp Clr Disch, {temperature}'].split()[column_location]
                temp_loc = column_location
                column_location += stages[f'Service {service + 1} Stage {int(stage + 1)}']
                throws=[cylinders[temp_loc][4].strip()]
                action=[cylinders[temp_loc][6]]
                ce_spacers = [cylinders[temp_loc][8]]
                he_spacers = [cylinders[temp_loc][7]]
                he_pocket = [cylinders[temp_loc][9]]
                he_pocket_used = [cylinders[temp_loc][10]]
                temp_loc += 1
                while temp_loc < column_location:
                    throws.append(cylinders[temp_loc][4].strip())
                    action.append(cylinders[temp_loc][6])
                    ce_spacers.append(cylinders[temp_loc][8])
                    he_spacers.append(cylinders[temp_loc][7])
                    he_pocket.append(cylinders[temp_loc][9])
                    he_pocket_used.append(cylinders[temp_loc][10])
                    stages[f'Service {service + 1} Stage {int(stage + 1)} Cylinder'].append(cylinders[temp_loc])
                    temp_loc += 1
                stages[f'Service {service + 1} Stage {int(stage + 1)} Throws'] = throws
                stages[f'Service {service + 1} Stage {int(stage + 1)} Action'] = action
                stages[f'Service {service + 1} Stage {int(stage + 1)} CE Spacers'] = ce_spacers
                stages[f'Service {service + 1} Stage {int(stage + 1)} HE Spacers'] = he_spacers
                stages[f'Service {service + 1} Stage {int(stage + 1)} HE Pocket'] = he_pocket
                stages[f'Service {service + 1} Stage {int(stage + 1)} HE Pocket Used'] = he_pocket_used

        # for stage in stages.items():
        #     print(stage)

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

        # def pressure_corrector(col_start, tot_cyl):
        #     ps_pd = ["", ""]
        #     if g_or_abs == 'Gauge':
        #         ps_pd[0] = round(float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]) + float(output_dict[f'Barmtr,{pressure[:1]+"a"}']), 3)
        #         pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
        #         ps_pd[1] = round(float(pd_lst[-1]) + float(output_dict[f'Barmtr,{pressure[:1]+"a"}']), 3)
        #         return ps_pd
        #     else:
        #         ps_pd[0] = round(float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]), 3)
        #         pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
        #         ps_pd[1] = round(float(pd_lst[-1]), 3)
        #         return ps_pd

        def pressure_corrector(col_start, tot_cyl):
            ps_pd = ["", ""]
            if g_or_abs == 'Gauge' and pressure != 'atm':
                ps_pd[0] = round(float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]) + float(output_dict[f'Barmtr,{pressure[:1]+"a"}']), 3)
                pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
                ps_pd[1] = round(float(pd_lst[-1]) + float(output_dict[f'Barmtr,{pressure[:1]+"a"}']), 3)

            elif pressure == 'atm':
                ps_pd[0] = round(float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]) + float(output_dict[f'Barmtr,{pressure}']), 3)
                pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
                ps_pd[1] = round(float(pd_lst[-1]) + float(output_dict[f'Barmtr,{pressure}']), 3)
            else:
                ps_pd[0] = round(float(output_dict[f'Pres Suct Line, {pressure}'].split()[col_start]), 3)
                pd_lst = [pd for pd in output_dict[f'Pres Disch Line, {pressure}'].split()[:col_start+tot_cyl] if pd != '---' and pd != 'N/A']
                ps_pd[1] = round(float(pd_lst[-1]), 3)

            if pressure_for_output == "psi":
                pass
            elif pressure_for_output == "bar":
                ps_pd = [round(p * 14.5038, 3) for p in ps_pd]
            elif pressure_for_output == "kPa":
                ps_pd = [round(p * 0.145038, 3) for p in ps_pd]
            elif pressure_for_output == "MPa":
                ps_pd = [round(p * 145.038, 3) for p in ps_pd]
            elif pressure_for_output == "kg/cm2":
                ps_pd = [round(p * 14.2233, 3) for p in ps_pd]
            elif pressure_for_output == "atm":
                ps_pd = [round(p * 14.6959, 3) for p in ps_pd]
            return ps_pd

        def mawp_corrector(press):
            if pressure_for_output == "psi":
                return press
            elif pressure_for_output == "bar":
                return int(5 * round((float(press) * 14.5038)/5))
            elif pressure_for_output == "kPa":
                return int(5 * round((float(press) * 0.145038)/5))
            elif pressure_for_output == "MPa":
                return int(5 * round((float(press) * 145.03774)/5)) #MPa cylinder MAWP in program might be a little off
            elif pressure_for_output == "kg/cm2":
                return int(5 * round((float(press) * 14.2233)/5))
            elif pressure_for_output == "atm":
                return int(5 * round((float(press) * 14.6959)/5))




        def cyl_assigner(total_cyls, stg, serv):
            cyl_output = ""
            # 0-model, 1-bore, 2-rdp, 3-mawp, 4-throw, 5-stage, 6-action, 7-he spcrs, 8-ce spcrs, 9-pkt avail, 10-pkt used
            for cyl in range(total_cyls):
                cyl_output = cyl_output + (f"""
                                <Cylinder>
                                    <throw>{stages[f'Service {serv} Stage {int(stg + 1)} Throws'][cyl]}</throw>
                                    <model>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][0]}</model>
                                    <mawp>{mawp_corrector(stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][3])}</mawp>
                                    <bore_size>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][1]}</bore_size>
                                    <action>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][6]}</action>
                                    <CESpacers>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][8]}</CESpacers>
                                    <HESpacers>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][7]}</HESpacers>
                                    <HEPocket>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][9]}</HEPocket>
                                    <HEPocketUsed>{stages[f'Service {serv} Stage {stg + 1} Cylinder'][cyl][10]}</HEPocketUsed>
                                </Cylinder>""")
            return cyl_output

        # The stage_assigner takes the service as an argument and populates an output
        #  string with each cylinder of that stage in the runM format.
        def stage_assigner(service):
            output = ""
            for stage in range(int(stages[f'Service {service} Total Stages'])):
                output = output + (f"""
                        <Stage>
                        	<number>{stage + 1}</number>
                            <suctionTemp>{stages[f'Service {service} Stage {stage + 1} Suction Temp']}</suctionTemp>
                            <coolerTemp>{cooler_temp_converter(stages[f'Service {service} Stage {stage + 1} Cooler Temp'])}</coolerTemp>
                        	<Cylinders>
                                <total>{stages[f'Service {service} Stage {stage + 1}']}</total>{cyl_assigner(int(stages[f'Service {service} Stage {stage + 1}']), stage, service)}
                        	</Cylinders>
                        </Stage>""")
            return output

        # temp_converter converts any temp to R.  Seems to consistently work. As of 9/1/21 Ariel7
        # expects suction temp in R but cooler temp in F.
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

        def cooler_temp_converter(cool_temp):
            if temperature =="F":
                temp_converted = cool_temp
            elif cool_temp == "N/A":
                temp_converted = cool_temp
            elif temperature == "C":
                temp_converted = round(float(cool_temp) * 9/5 + 32, 2)
            elif temperature == "K":
                temp_converted = round((float(cool_temp) -273.15)* 9/5 + 32, 2)
            else:
                temp_converted = round(float(cool_temp) - 459.67 , 2)
            return temp_converted


        # Converts elevation to meters if required
        def elevation_converter(elevation):
            if eng_met == 'English':
                return elevation
            else:
                return round(float(elevation) * 3.28084, 3)

        # gas_type looks at the sg and takes a guess at the gas_type based on that.
        #  Searches the text string for the phrase 'Sour Gas' and if it is found
        # the gas type is set to sour gas, otherwise the gas type is set based
        # on a guess at the SG.
        def gas_type(sg):
            sgf = float(sg)
            if sgf < 1.4500 and sgf > 1.2000 and re.search('SOUR GAS', text):
                return 'Sour Gas'
            elif sgf >= 0.6000 and sgf <= 1.1000:
                return 'Field Gas'
            elif sgf > 1.1000 and sgf <=2.0000:
                return 'Heavy Hydrocarbons'
            elif sgf >= 0.5500 and sgf <0.6000:
                return 'Pipeline Quality'


        # volume_converter converts any volume units to MMSCFD.  As of this writing
        # Ariel 7 does not set the correct volume units when provided kg/h or lb/h
        def volume_converter(vol_type, vol_num):
            if vol_type == "MMSCFD":
                return vol_num
            elif vol_type == "MSCFD":
                return round(float(vol_num) * 0.001, 3)
            elif vol_type == "SCFH":
                return round(float(vol_num) * 0.000024, 3)
            elif vol_type == "SCFM":
                return round(float(vol_num) * 0.00144, 3)
            elif vol_type == "lb/h":
                return round(float(vol_num) / 2066.953, 3)
            elif vol_type == "Nm3/h":
                return round(float(vol_num) / 1116.641, 3)
            elif vol_type == "Sm3/h":
                return round(float(vol_num) / 1177.961, 3)
            elif vol_type == "E3m3/D":
                return round(float(vol_num) / 28.271, 3)
            elif vol_type == "kg/h":
                return round(float(vol_num) / 937.555, 3)


        # service_assigner loops through each service of the machine and populates an
        #  output string with the required information in runM format.  it calls the
        #  stage_assigner function for each service to populate the stages portion and
        #  calls some of the other functions to populate the particular value in question.
        def service_assigner():
            output = ""
            for service in range(1, int(stages[f'Total Services']) + 1):
                output = output + f"""
                <Service>
                    <target_flow>{volume_converter(flow, output_dict[f'Target Flow, {flow}'].split()[stages[f'Service {service} Column Start']])}</target_flow>
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
                    <Stages>{stage_assigner(service)}
                    </Stages>
                </Service>"""
            return output


        # Finally the program creates an output_txt string which creates the top portion
        #  of the runM format and then calls the service assigner to fill in service
        #  and stage specific information.

        output_txt = fr"""<?xml version="1.0" encoding=UTF-8 ?>
        <Ariel_Mobile>
            <Type>PDF</Type>
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
        		<elevation>{elevation_converter(output_dict["Elevation"])}</elevation>
        	</User>
        	<Compressor>
        		<RPM>{output_dict['Calc RPM']}</RPM>
        		<model>{output_dict['Frame'].replace("(ELP)", "")[:-2].strip()}</model>
        		<num_throws>{output_dict['Frame'][-1]}</num_throws>
                <Driver>
                    <type>{output_dict['Driver Type'].strip()}</type>
                    <mfg>{output_dict['Driver Mfg'].strip()}</mfg>
                    <model>{output_dict['Driver Model'].strip()}</model>
                    <ratedBHP>{driver_rated_BHP}</ratedBHP>
                    <derate>{driver_derate}</derate>
                    <aux>{driver_aux}</aux>
                </Driver>
        	</Compressor>
            <Services>{service_assigner()}
        	</Services>
        </Ariel_Mobile>
        """

        print(output_txt)



        with open(save_filepath, mode='w', encoding='utf-8') as file:
            file.writelines(output_txt)

        os.startfile(save_filepath)
    except Exception as e:
        error_box("Program Error", e)

###################################################
###################################################
###################################################
###################################################

#This is the start of the GUI portion.  All of the code up to here is saved in
# the run_calcs function so none of it will run until the take_ent() function
# is activated by clicking the next button.

#open_file and save_file open dialogue boxes that direct you to choose a source
# .pdf and the save location for the resultant runM.  The variable are made
# global so that they could be seen outside of the GUI namespace.  There is probably
# another way around this but these two don't pollute the global name space enough
# to cause an issue.
def open_file():
    global open_filepath
    open_filepath = askopenfilename(filetypes=[('PDF', '*.pdf')])
    open_var.set(open_filepath)
    if not open_filepath:
        return

def save_file():
    global save_filepath
    save_filepath = asksaveasfilename(defaultextension='runM', filetypes=[("runM", "*.runM")])
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
    run_calcs()




window = tk.Tk()
window.title("Auto Run File Builder")
window.geometry("900x400")
window.minsize(900, 400)

#open source file location
frm_1 = tk.Frame(master = window)
frm_1.grid(row=0, column = 0, padx=20, pady=(100,20), sticky='w')
# open_file_label = tk.Label(master=frm_1, text="Select a source PDF run report")
# open_file_label.grid(row=0, columnspan = 2, sticky = tk.W+tk.E)
btn_open = tk.Button(master = frm_1, text = "Select source file", width=20, command=open_file)
btn_open.grid(row=0, column=0)

open_var = tk.StringVar()
lbl_open = tk.Label(master=frm_1, textvariable=open_var, width=70, relief=tk.SUNKEN, borderwidth=2)
lbl_open.grid(row=0, column=1, padx=20)

#save result file location
frm_2 = tk.Frame(master=window)
frm_2.grid(row = 12, column = 0, padx = 20, pady = (50,50), sticky = 'w')
btn_save = tk.Button(master = frm_2, text='Results save location', width=20, command=save_file)
btn_save.grid(row=0, column=0)

save_var = tk.StringVar()
lbl_save = tk.Label(master=frm_2, textvariable=save_var, width=70, relief=tk.SUNKEN, borderwidth=2)
lbl_save.grid(row=0, column=1, padx=20)

# Next Button
btn = tk.Button(master=window, text="Next", width = 15, height = 3, command = take_ent)
btn.grid(row=15, column=1, sticky=tk.S)


window.mainloop()

# # End of GUI portion
# ######################################################################
