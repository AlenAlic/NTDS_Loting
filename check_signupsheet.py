from tkinter import *
import textwrap
from classes.entrybox import EntryBox
import os
import openpyxl
from loting import total_width, status_text_width, xlsx_ext, roles, levels, options_yn, gen_dict

CODES = {'error': 'ERROR: '}

file_key = {'name': 'Test_Enschede.xlsx', 'path': 'test_control'}
file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['path'])
file_key['path'] = os.path.join(file_key['path'], file_key['name'])


def wip():
    """"Temp"""
    status_print('Work in progress...')


def column(matrix, i):
    """"Temp"""
    return [row[i] for row in matrix]


def status_print(message, wrap=True, code=None):
    """"Prints the message passed to the program screen"""
    if code in list(CODES.values()):
        if code == CODES['error']:
            message = code + message
    status_text.config(state=NORMAL)
    if wrap is True:
        message = textwrap.fill(message, status_text_width)
    status_text.insert(END, message)
    status_text.insert(END, '\n')
    status_text.update()
    status_text.see(END)
    status_text.config(state=DISABLED)


def welcome_text():
    """"Text displayed when opening the program for the first time."""
    status_text.config(state=NORMAL)
    status_text.config(wrap=WORD)
    status_text.config(state=DISABLED)
    status_print('Welcome to the NTDS Signup Sheet Check-O-Matic 3000.')
    status_print('')
    status_print('This program will check for errors in a filled in signup sheet. '
                 'Any found errors will be displayed in window to the right.')
    status_print('')
    status_print('Please select a file to check for errors, using the button in the bottom right corner '
                 '(you know wich one).')
    status_print('')
    status_print('DISCLAIMER:')
    status_print('This program can only check for system errors, for example two Followers that signed up as partners '
                 'together. Registering into the Open Class instead of the Breitensport by accident, '
                 'and similar errors will NOT be flagged.')
    status_print('')
    status_text.config(state=NORMAL)
    status_text.config(wrap=NONE)
    status_text.config(state=DISABLED)
    data_text.config(state=NORMAL)
    data_text.delete('1.0', END)
    data_text.insert(END, 'Any errors found after checking a file will be displayed here.')
    data_text.config(state=DISABLED)


def select_file():
    """"Temp"""
    if file_key['name'].endswith(xlsx_ext):
        old_name = file_key['name']
    else:
        old_name = file_key['name'] + xlsx_ext
    ask_database = EntryBox('Please give the file name', (file_key, 'name'))
    root.wait_window(ask_database.top)
    if file_key['name'].endswith(xlsx_ext):
        file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['name'])
    else:
        file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['name'] + xlsx_ext)
    if os.path.isfile(path=file_key['path']):
        status_text.config(state=NORMAL)
        status_text.delete('1.0', END)
        status_text.config(state=DISABLED)
        status_print('Selected existing database: {path}'.format(path=file_key['path']))
    elif os.path.isfile(path=file_key['path']) is False or old_name.lower() != file_key['name'].lower():
        status_print('The file "{name}" does not exist.'.format(name=file_key['name']))


def check_file():
    """Temp"""
    if os.path.isfile(path=file_key['path']):
        status_print('Checking file: {file}'.format(file=file_key['path']))
        status_print('')
        status_print('')
        workbook = openpyxl.load_workbook(file_key['path'], data_only=True)
        worksheet = workbook.worksheets[0]
        signup_list = list(worksheet.iter_rows(min_col=1, min_row=2, max_col=22, max_row=201))
        # Convert data to 2d list, replace None values with an empty string,
        # increase the id numbers so that there are no duplicates, and add the city to the contestant
        signup_list = [[cell.value for cell in row] for row in signup_list]
        signup_list = [['' if elem is None else elem for elem in row] for row in signup_list]
        for index, row in enumerate(signup_list):
            if row[gen_dict['id']] == '':
                signup_list = signup_list[0:index]
                break
        # Check number of team captains
        status_print('Checking number of team captains.')
        status_print('')
        team_captains = column(signup_list, gen_dict['team_captain'])
        team_captain_ids = [i + 1 for i, x in enumerate(team_captains) if x == options_yn['yes']]
        for tc in team_captain_ids:
            status_print('Contestant number {num} signed up as a team captain.'.format(num=tc))
        number_of_team_captains = team_captains.count(options_yn['yes'])
        status_print('Number of team captains is: {num}.'.format(num=number_of_team_captains))
        if number_of_team_captains == 0:
            status_print('Missing at least one team captain.')
        elif 0 < number_of_team_captains <= 2:
            status_print('Number of team captains is OK, continuing check.')
        elif number_of_team_captains > 2:
            status_print('Too many team captains have signed in. The maximum is two.',
                         code=CODES['error'])
        status_print('')
        status_print('')
        # TODO Check if partner numbers are correct
        status_print('Checking each individual contestant.')
        status_print('')
        for row in signup_list:
            contestant_number = row[gen_dict['id']]
            status_print('Checking contestant {num}.'.format(num=contestant_number))
            contestant_ballroom_level = row[gen_dict['ballroom_level']]
            contestant_ballroom_partner_number = row[gen_dict['ballroom_partner']]
            contestant_ballroom_role = row[gen_dict['ballroom_role']]
            contestant_ballroom_bd = row[gen_dict['ballroom_mandatory_blind_date']]
            contestant_team_captain = row[gen_dict['team_captain']]
            contestant_latin_level = row[gen_dict['latin_level']]
            contestant_latin_partner_number = row[gen_dict['latin_partner']]
            contestant_latin_role = row[gen_dict['latin_role']]
            contestant_latin_bd = row[gen_dict['latin_mandatory_blind_date']]
            if all([contestant_ballroom_level == '', contestant_latin_level == '', contestant_ballroom_role == '',
                    contestant_latin_role == '', contestant_team_captain != options_yn['yes']]):
                status_print('Contestant number {contestant} is not participating in any category, and is not a '
                             'team captain.'.format(contestant=contestant_number))
            if contestant_ballroom_partner_number != '' and contestant_latin_partner_number != '':
                partner_row = signup_list[contestant_ballroom_partner_number - 1]
                second_partner_row = signup_list[contestant_latin_partner_number - 1]
                partner_number = partner_row[gen_dict['id']]
                partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
                partner_latin_number = partner_row[gen_dict['latin_partner']]
                partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
                partner_latin_role = partner_row[gen_dict['latin_role']]
                partner_ballroom_bd = partner_row[gen_dict['ballroom_mandatory_blind_date']]
                partner_latin_bd = partner_row[gen_dict['latin_mandatory_blind_date']]
                if partner_row == second_partner_row:
                    status_print('Contestant number {contestant} signed up for both Ballroom and Latin with contestant '
                                 'number {partner}.'.format(contestant=contestant_number, partner=partner_number))
                    if partner_ballroom_number == contestant_number and partner_latin_number == contestant_number:
                        status_print('Contestant number {partner} signed up for both Ballroom and Latin with '
                                     'contestant number {contestant} as well.'
                                     .format(contestant=contestant_number, partner=partner_number))
                    else:
                        status_print('Contestant number {partner} did not sign up for both Ballroom and Latin with '
                                     'contestant number {contestant}.'
                                     .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                    if all([contestant_ballroom_role != '', contestant_latin_role != '',
                            partner_ballroom_role != '', partner_latin_role != '',
                            contestant_ballroom_role != partner_ballroom_role,
                            contestant_latin_role != partner_latin_role]):
                        status_print('Contestants number {contestant} and {partner} have opposite roles in both '
                                     'Ballroom and Latin categories.'
                                     .format(contestant=contestant_number, partner=partner_number))
                    else:
                        status_print('Contestants number {contestant} and {partner} have matching roles in either the '
                                     'Ballroom or Latin categories.'
                                     .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                    if all([contestant_ballroom_bd != options_yn['yes'], contestant_latin_bd != options_yn['yes'],
                            partner_ballroom_bd != options_yn['yes'], partner_latin_bd != options_yn['yes']]):
                        status_print('Neither contestant number {contestant} or {partner} have to Blind Date.'
                                     .format(contestant=contestant_number, partner=partner_number))
                    else:
                        status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                                     'Blind Date, so these contestants are not allowed to dance together.'
                                     .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                else:
                    # TODO Check in case of different partners
                    second_partner_number = second_partner_row[gen_dict['id']]
                    second_partner_latin_number = second_partner_row[gen_dict['latin_partner']]
                    second_partner_latin_role = second_partner_row[gen_dict['latin_role']]
                    second_partner_latin_bd = second_partner_row[gen_dict['latin_mandatory_blind_date']]
                    status_print('Contestant number {contestant} signed up Ballroom with contestant number {partner} '
                                 'and for Latin with contestant number {second_partner}.'
                                 .format(contestant=contestant_number, partner=partner_number,
                                         second_partner=second_partner_number))
            elif contestant_ballroom_partner_number == '' and contestant_latin_partner_number == '':
                status_print('Contestant number {contestant} signed up as a Blind Dater in both Ballroom and Latin.'
                             .format(contestant=contestant_number))
                if any([all([contestant_ballroom_level != '', contestant_latin_level != '',
                             contestant_ballroom_role != '', contestant_latin_role != '']),
                        all([contestant_ballroom_level != '', contestant_latin_level == '',
                             contestant_ballroom_role != '', contestant_latin_role == '']),
                        all([contestant_ballroom_level == '', contestant_latin_level != '',
                             contestant_ballroom_role == '', contestant_latin_role != ''])]):
                    status_print('Signup data of contestant number {contestant} is OK.'
                                 .format(contestant=contestant_number))
                else:
                    status_print('contestant alone NOT OK')
            # TODO Check when only one partner is signed
            elif contestant_ballroom_partner_number != '' and contestant_latin_partner_number == '':
                status_print('only ballroom partner')
            elif contestant_ballroom_partner_number == '' and contestant_latin_partner_number != '':
                status_print('only latin partner')
            # if all([contestant_ballroom_level == '', contestant_latin_level == '', contestant_ballroom_role == '',
            #         contestant_latin_role == '', contestant_team_captain != options_yn['yes']]):
            #     status_print('Contestant number {num} is not participating in any category, and is not a team captain.'
            #                  .format(num=contestant_number))
            # else:
            #     if contestant_ballroom_level != '':
            #         status_print('Contestant number {num} is participating in the {level} Ballroom category.'
            #                      .format(num=contestant_number, level=contestant_ballroom_level))
            #         if contestant_ballroom_role in list(roles.values()):
            #             status_print('Contestant number {num} is dancing as a {role} in the Ballroom category.'
            #                          .format(num=contestant_number, role=contestant_ballroom_role))
            #             # if contestant_ballroom_role == levels['open_class']:
            #             #     if contestant_ballroom_bd != options_yn['yes']:
            #             #         status_print(
            #             #             'Contestant number {num} is participating in the {level} Ballroom category, but '
            #             #             .format(num=contestant_number, level=contestant_ballroom_level), code=CODES['error'])
            #         else:
            #             status_print('Contestant number {num} has not given a role in the Ballroom category.'
            #                          .format(num=contestant_number))
            #
            #         if isinstance(contestant_ballroom_partner_number, int):
            #             status_print('Contestant number {num} signed op with contestant number {partner} for the Ballroom '
            #                          'category.'.format(num=contestant_number, partner=contestant_ballroom_partner_number))
            #             partner_row = signup_list[contestant_ballroom_partner_number - 1]
            #             partner_number = partner_row[gen_dict['id']]
            #             partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
            #             if isinstance(partner_ballroom_number, int):
            #                 if partner_ballroom_number == contestant_number:
            #                     status_print('Contestant number {partner} signed op with contestant number {num} for the '
            #                                  'Ballroom category.'
            #                                  .format(num=contestant_number, partner=partner_number))
            #                 else:
            #                     status_print('Contestant number {partner} signed op with contestant number {num} for the '
            #                                  'Ballroom category.'
            #                                  .format(num=partner_ballroom_number, partner=partner_number),
            #                                  code=CODES['error'])
            #             else:
            #                 status_print('Contestant number {partner} signed op without a partner for the Ballroom '
            #                              'category.'
            #                              .format(partner=contestant_ballroom_partner_number), code=CODES['error'])
            #             partner_ballroom_level = partner_row[gen_dict['ballroom_level']]
            #
            #             partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
            #             if contestant_ballroom_role != '':
            #                 if contestant_ballroom_role == partner_ballroom_role:
            #                     status_print('Contestant number {num} is and contestant number {partner} are both dancing '
            #                                  'as a {contestant_role} in the Ballroom category.'
            #                                  .format(num=contestant_number, contestant_role=contestant_ballroom_role,
            #                                          partner=partner_number))
            #                 if contestant_ballroom_role != partner_ballroom_role and partner_ballroom_role != '':
            #                     status_print('Contestant number {num} is dancing as a {contestant_role} and contestant '
            #                                  'number {partner} is dancing as a {partner_role} in the Ballroom category.'
            #                                  .format(num=contestant_number, contestant_role=contestant_ballroom_role,
            #                                          partner=partner_number, partner_role=partner_ballroom_role))
            #             else:
            #                 status_print('Contestant number {num} has not given a role for for the Ballroom category.'
            #                              .format(num=contestant_number), code=CODES['error'])
            #             # if partner_ballroom_number != contestant_number:
            #             #     if isinstance(partner_ballroom_number, int):
            #             #         status_print('In Ballroom, contestant number {contestant} has signed up with contestant '
            #             #                      'number {ballroom_partner}, while {ballroom_partner} signed up with '
            #             #                      'contestant number {partner_ballroom_number}.'
            #             #                      .format(contestant=contestant_number, ballroom_partner=partner_number,
            #             #                              partner_ballroom_number=partner_ballroom_number),
            #             #                      code=CODES['error'])
            #             #     else:
            #             #         status_print('In Ballroom, contestant number {contestant} has signed up with contestant '
            #             #                      'number {ballroom_partner}, while {ballroom_partner} did not.'
            #             #                      .format(contestant=contestant_number, ballroom_partner=partner_number),
            #             #                      code=CODES['error'])
            #             # else:
            #             #     status_print('In Ballroom, contestants number {contestant} and {ballroom_partner} signed up '
            #             #                  'together.'.format(contestant=contestant_number, ballroom_partner=partner_number))
            #             # if contestant_ballroom_role == roles['lead']:
            #             #     if partner_ballroom_role == roles['lead']:
            #             #         status_print('In Ballroom, contestant number {contestant} and contestant number '
            #             #                      '{ballroom_partner} both have a selected their Ballroom role as {lead}.'
            #             #                      .format(contestant=contestant_number, ballroom_partner=partner_number,
            #             #                              lead=roles['lead']), code=CODES['error'])
            #             #     elif partner_ballroom_role == roles['follow']:
            #             #         status_print('In Ballroom, contestant number {contestant} dances as a {contestant_role} '
            #             #                      'and contestant number {ballroom_partner} dances as a {partner_role}.'
            #             #                      .format(contestant=contestant_number, contestant_role=contestant_ballroom_role,
            #             #                              ballroom_partner=partner_number, partner_role=partner_ballroom_role))
            #             # if contestant_ballroom_role == roles['follow']:
            #             #     if partner_ballroom_role == roles['follow']:
            #             #         status_print('In Ballroom, contestant number {contestant} and contestant number '
            #             #                      '{ballroom_partner} both have a selected their Ballroom role as {lead}.'
            #             #                      .format(contestant=contestant_number, ballroom_partner=partner_number,
            #             #                              lead=roles['lead']), code=CODES['error'])
            #             #     elif partner_ballroom_role == roles['lead']:
            #             #         status_print('In Ballroom, contestant number {contestant} dances as a {contestant_role} '
            #             #                      'and contestant number {ballroom_partner} dances as a {partner_role}.'
            #             #                      .format(contestant=contestant_number, contestant_role=contestant_ballroom_role,
            #             #                              ballroom_partner=partner_number, partner_role=partner_ballroom_role))
            #         if contestant_ballroom_partner_number == '':
            #             status_print('Contestant number {contestant} signed up without a partner in the Ballroom category.'
            #                          .format(contestant=contestant_number))
            #             if contestant_ballroom_bd == options_yn['yes']:
            #                 status_print('Contestant number {contestant} signed up as a mandatory Blind Dater in the '
            #                              'Ballroom category.'.format(contestant=contestant_number))
            #             else:
            #                 status_print('Contestant number {contestant} signed up as a voluntary Blind Dater in the '
            #                              'Ballroom category.'.format(contestant=contestant_number))
            #     else:
            #         status_print('Contestant {num} is not participating in the Ballroom category.'
            #                      .format(num=contestant_number))
            # # if contestant_latin_level != '':
            # #     if isinstance(contestant_latin_partner_number, int):
            # #         partner_row = signup_list[contestant_latin_partner_number - 1]
            # #         partner_number = partner_row[gen_dict['id']]
            # #         partner_latin_number = partner_row[gen_dict['latin_partner']]
            # #         if partner_latin_number != contestant_number:
            # #             if isinstance(partner_latin_number, int):
            # #                 status_print('In Latin, Contestant number {contestant} has signed up with contestant '
            # #                              'number {ballroom_partner}, while {ballroom_partner} signed up with '
            # #                              'contestant number {partner_ballroom_number}.'
            # #                              .format(contestant=contestant_number, ballroom_partner=partner_number,
            # #                                      partner_ballroom_number=partner_latin_number),
            # #                              code=CODES['error'])
            # #             else:
            # #                 status_print('In Latin, Contestant number {contestant} has signed up with contestant '
            # #                              'number  {ballroom_partner}, while {ballroom_partner} did not.'
            # #                              .format(contestant=contestant_number, ballroom_partner=partner_number),
            # #                              code=CODES['error'])
            # #         else:
            # #             status_print('In Latin, contestants number {contestant} and {ballroom_partner} signed up '
            # #                          'together.'.format(contestant=contestant_number, ballroom_partner=partner_number))
            # #     elif contestant_latin_partner_number == '':
            # #         status_print('In Latin, contestant number {contestant} signed up alone.'
            # #                      .format(contestant=contestant_number))
            # # else:
            # #     status_print('Contestant {num} is not participating in the Latin category.'
            # #                  .format(num=contestant_number))
            # TODO Check if beginners signed up as opbligatory blind daters warming
            # TODO Niet open klasse die wil jureren error
            # TODO open klasse die niet verplicht moet blind date warning
            status_print('Checking contestant {num} done. Continuing check.'.format(num=contestant_number))
            status_print('')
        # for row in signup_list:
        #     contestant = row[0] + 1
        #     specific_volunteer = row[16:19]
        #     volunteer = row[14]
        #     if volunteer == 'Misschien' and specific_volunteer.count('Ja') > 0:
        #         status_print('soft warning')
        #     elif volunteer == 'Nee' and specific_volunteer.count('Ja') > 0:
        #         status_print('WARNING')
        #     elif volunteer == 'Nee' and specific_volunteer.count('Misschien') > 0:
        #         status_print('WARNING')
        #     else:
        #         status_print('Volunteering work of contestant number {num} is OK, continuing check.'
        #                      .format(num=contestant))
    elif os.path.isfile(path=file_key['path']) is False:
        status_print('The file "{name}" does not exist.'.format(name=file_key['name']))


if __name__ == "__main__":
    root = Tk()
    root.geometry("1600x900")
    root.state('zoomed')
    root.title('NTDS Signup Sheet Check-O-Matic 3000')
    pad_out = 8
    pad_in = 8
    frame = Frame()
    frame.place(in_=root, anchor="c", relx=.50, rely=.50)
    x_scrollbar = Scrollbar(master=frame, orient=HORIZONTAL)
    x_scrollbar.grid(row=1, column=0, padx=pad_in, sticky=E+W)
    y_scrollbar = Scrollbar(master=frame, orient=VERTICAL)
    y_scrollbar.grid(row=0, column=1, pady=pad_in, sticky=N+S)
    status_text = Text(master=frame, width=status_text_width, height=50, padx=pad_in, pady=pad_in,
                       xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set, state=DISABLED, wrap=NONE)
    status_text.grid(row=0, column=0, padx=pad_out)
    x_scrollbar.config(command=status_text.xview)
    y_scrollbar.config(command=status_text.yview)
    pad_frame = Frame(master=frame, height=16)
    pad_frame.grid(row=2, column=0, padx=pad_in, pady=pad_out)
    data_help_frame = Frame(master=frame)
    data_help_frame.grid(row=0, column=2, rowspan=3, columnspan=3)
    data_text = Text(master=data_help_frame, width=total_width - status_text_width, height=50,
                     padx=pad_in, pady=pad_in, wrap=WORD, state=DISABLED)
    data_text.grid(row=0, column=0, padx=pad_out, columnspan=3)
    padding_frame = Frame(master=data_help_frame, height=16)
    padding_frame.grid(row=1, column=0)
    start_button = Button(master=data_help_frame, text='Select file', command=select_file)
    start_button.grid(row=2, column=0, padx=pad_out, pady=pad_in)
    update_button = Button(master=data_help_frame, text='Check file', command=check_file)
    update_button.grid(row=2, column=1, padx=pad_out)
    welcome_text()
    root.mainloop()
