from tkinter import *
import textwrap
from classes.entrybox import EntryBox
import os
import openpyxl
from loting import total_width, status_text_width, xlsx_ext, levels, options_yn, options_ymn, gen_dict

CODES = {'error': 'ERROR: ', 'warning': 'WARNING: '}

file_key = {'name': '', 'path': ''}
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
    status_print('')
    status_print('This program will check for errors in a filled in signup sheet. '
                 'Any found errors will be displayed here.')
    status_print('')
    status_print('Please select a file to check for errors, using the button below (you know wich one).')
    status_print('')
    status_print('DISCLAIMER:')
    status_print('This program can only check for system errors, for example two Followers that signed up as partners '
                 'together. Registering into the Open Class instead of the Breitensport by accident, '
                 'and similar errors will NOT be flagged.')
    status_print('')
    status_text.config(state=NORMAL)
    status_text.config(wrap=NONE)
    status_text.config(state=DISABLED)


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


def check_contestants(contestants_list):
    """"Temp"""
    for row in contestants_list:
        contestant_number = row[gen_dict['id']]
        status_print('Checking contestant {num}.'.format(num=contestant_number))
        contestant_ballroom_level = row[gen_dict['ballroom_level']]
        contestant_ballroom_partner_number = row[gen_dict['ballroom_partner']]
        contestant_ballroom_role = row[gen_dict['ballroom_role']]
        contestant_ballroom_bd = row[gen_dict['ballroom_mandatory_blind_date']]
        contestant_latin_level = row[gen_dict['latin_level']]
        contestant_latin_partner_number = row[gen_dict['latin_partner']]
        contestant_latin_role = row[gen_dict['latin_role']]
        contestant_latin_bd = row[gen_dict['latin_mandatory_blind_date']]
        # Contestant signed up alone for both categories without a partner
        if contestant_ballroom_partner_number == '' and contestant_latin_partner_number == '':
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
        # Contestant signed up with a partner for both Ballroom and Latin
        elif contestant_ballroom_partner_number != '' and contestant_latin_partner_number != '':
            partner_row = contestants_list[contestant_ballroom_partner_number - 1]
            second_partner_row = contestants_list[contestant_latin_partner_number - 1]
            partner_number = partner_row[gen_dict['id']]
            partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
            partner_latin_number = partner_row[gen_dict['latin_partner']]
            partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
            partner_latin_role = partner_row[gen_dict['latin_role']]
            partner_ballroom_bd = partner_row[gen_dict['ballroom_mandatory_blind_date']]
            partner_latin_bd = partner_row[gen_dict['latin_mandatory_blind_date']]
            # Contestant signed up with the same partner for both Ballroom and Latin
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
                    status_print('Contestants number {contestant} and {partner} don\'t have to Blind Date, '
                                 'and are allowed to dance together.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                                 'Blind Date, so these contestants are not allowed to dance together.'
                                 .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
            # Contestant signed up with the different partners for Ballroom and Latin
            else:
                second_partner_number = second_partner_row[gen_dict['id']]
                second_partner_latin_number = second_partner_row[gen_dict['latin_partner']]
                second_partner_latin_role = second_partner_row[gen_dict['latin_role']]
                second_partner_latin_bd = second_partner_row[gen_dict['latin_mandatory_blind_date']]
                status_print('Contestant number {contestant} signed up for Ballroom with contestant number '
                             '{partner} and for Latin with contestant number {second_partner}.'
                             .format(contestant=contestant_number, partner=partner_number,
                                     second_partner=second_partner_number))
                if partner_ballroom_number == contestant_number:
                    status_print('Contestant number {partner} signed up for Ballroom with contestant number '
                                 '{contestant} as well.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestant number {partner} did not sign up for Ballroom with contestant number '
                                 '{contestant}.'
                                 .format(contestant=contestant_number, partner=partner_number),
                                 code=CODES['error'])
                if second_partner_latin_number == contestant_number:
                    status_print('Contestant number {partner} signed up for Latin with contestant number '
                                 '{contestant} as well.'
                                 .format(contestant=contestant_number, partner=second_partner_number))
                else:
                    status_print('Contestant number {partner} did not sign up for Latin with contestant number '
                                 '{contestant}.'
                                 .format(contestant=contestant_number, partner=second_partner_number),
                                 code=CODES['error'])
                if all([contestant_ballroom_role != '', partner_ballroom_role != '',
                        contestant_ballroom_role != partner_ballroom_role]):
                    status_print('Contestants number {contestant} and {partner} have opposite roles in the '
                                 'Ballroom category.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestants number {contestant} and {partner} have matching roles in the '
                                 'Ballroom category.'
                                 .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                if all([contestant_latin_role != '', second_partner_latin_role != '',
                        contestant_latin_role != second_partner_latin_role]):
                    status_print('Contestants number {contestant} and {partner} have opposite roles in the '
                                 'Latin category.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestants number {contestant} and {partner} have matching roles in the '
                                 'Latin category.'
                                 .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                if contestant_ballroom_bd != options_yn['yes'] and partner_ballroom_bd != options_yn['yes']:
                    status_print('Contestants number {contestant} and {partner} don\'t have to Blind Date, '
                                 'and are allowed to dance together in the Ballroom category.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                                 'Blind Date, so these contestants are not allowed to dance together in the '
                                 'Ballroom category.'
                                 .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
                if contestant_latin_bd != options_yn['yes'] and second_partner_latin_bd != options_yn['yes']:
                    status_print('Contestants number {contestant} and {partner} don\'t have to Blind Date, '
                                 'and are allowed to dance together in the Latin category.'
                                 .format(contestant=contestant_number, partner=partner_number))
                else:
                    status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                                 'Blind Date, so these contestants are not allowed to dance together in the '
                                 'Latin category.'
                                 .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
        # Contestant signed up dancing only one Ballroom with a partner
        elif contestant_ballroom_partner_number != '' and contestant_latin_partner_number == '':
            partner_row = contestants_list[contestant_ballroom_partner_number - 1]
            partner_number = partner_row[gen_dict['id']]
            partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
            partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
            partner_ballroom_bd = partner_row[gen_dict['ballroom_mandatory_blind_date']]
            status_print('Contestant number {contestant} signed up for Ballroom with contestant number {partner}.'
                         .format(contestant=contestant_number, partner=partner_number))
            if partner_ballroom_number == contestant_number:
                status_print('Contestant number {partner} signed up for Ballroom with contestant number '
                             '{contestant} as well.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestant number {partner} did not sign up for Ballroom with contestant number '
                             '{contestant}.'
                             .format(contestant=contestant_number, partner=partner_number),
                             code=CODES['error'])
            if all([contestant_ballroom_role != '', partner_ballroom_role != '',
                    contestant_ballroom_role != partner_ballroom_role]):
                status_print('Contestants number {contestant} and {partner} have opposite roles in the '
                             'Ballroom category.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestants number {contestant} and {partner} have matching roles in the '
                             'Ballroom category.'
                             .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
            if contestant_ballroom_bd != options_yn['yes'] and partner_ballroom_bd != options_yn['yes']:
                status_print('Contestants number {contestant} and {partner} don\'t have to Blind Date, '
                             'and are allowed to dance together in the Ballroom category.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                             'Blind Date, so these contestants are not allowed to dance together in the '
                             'Ballroom category.'
                             .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
        # Contestant signed up dancing only one Latin with a partner
        elif contestant_ballroom_partner_number == '' and contestant_latin_partner_number != '':
            partner_row = contestants_list[contestant_latin_partner_number - 1]
            partner_number = partner_row[gen_dict['id']]
            partner_latin_number = partner_row[gen_dict['latin_partner']]
            partner_latin_role = partner_row[gen_dict['latin_role']]
            partner_latin_bd = partner_row[gen_dict['latin_mandatory_blind_date']]
            status_print('Contestant number {contestant} signed up for Latin with contestant number {partner}.'
                         .format(contestant=contestant_number, partner=partner_number))
            if partner_latin_number == contestant_number:
                status_print('Contestant number {partner} signed up for Latin with contestant number '
                             '{contestant} as well.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestant number {partner} did not sign up for Latin with contestant number '
                             '{contestant}.'
                             .format(contestant=contestant_number, partner=partner_number),
                             code=CODES['error'])
            if all([contestant_ballroom_role != '', partner_latin_role != '',
                    contestant_ballroom_role != partner_latin_role]):
                status_print('Contestants number {contestant} and {partner} have opposite roles in the '
                             'Latin category.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestants number {contestant} and {partner} have matching roles in the '
                             'Latin category.'
                             .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
            if contestant_ballroom_bd != options_yn['yes'] and partner_latin_bd != options_yn['yes']:
                status_print('Contestants number {contestant} and {partner} don\'t have to Blind Date, '
                             'and are allowed to dance together in the Latin category.'
                             .format(contestant=contestant_number, partner=partner_number))
            else:
                status_print('Contestant number {contestant} or {partner} have indicated that they have to '
                             'Blind Date, so these contestants are not allowed to dance together in the '
                             'Latin category.'
                             .format(contestant=contestant_number, partner=partner_number), code=CODES['error'])
        # Check if contestant is a non dancing team captain
        contestant_team_captain = row[gen_dict['team_captain']]
        if all([contestant_ballroom_level == '', contestant_latin_level == '', contestant_ballroom_role == '',
                contestant_latin_role == '', contestant_team_captain != options_yn['yes']]):
            status_print('Contestant number {contestant} is not participating in any category, and is not a '
                         'team captain.'.format(contestant=contestant_number))
        # Check if a Beginner has accidentally indicated that he/she needs to Blind Date
        if contestant_ballroom_level == levels['beginners'] and contestant_ballroom_bd == options_yn['yes']:
            status_print('Contestant number {contestant} is a Beginner, but indicated that he/she is mandatory '
                         'Blind Dating.'.format(contestant=contestant_number), code=CODES['warning'])
        # Check if a non-Open Class dances has indicated that he/she wants to be a Jury
        contestant_ballroom_jury = row[gen_dict['ballroom_jury']]
        if all([contestant_ballroom_level != levels['open_class'], contestant_ballroom_level != '']) and \
                (contestant_ballroom_jury == options_ymn['yes'] or contestant_ballroom_jury ==
                    options_ymn['maybe']):
            status_print('Contestant number {contestant} is not an Open Class dancer, but indicated that he/she is '
                         'willing to be a jury in the Ballroom category.'
                         .format(contestant=contestant_number), code=CODES['error'])
        contestant_latin_jury = row[gen_dict['latin_jury']]
        if all([contestant_latin_level != levels['open_class'], contestant_latin_level != '']) and \
                (contestant_latin_jury == options_ymn['yes'] or contestant_latin_jury == options_ymn['maybe']):
            status_print('Contestant number {contestant} is not an Open Class dancer, but indicated that he/she is '
                         'willing to be a jury in the Latin category.'
                         .format(contestant=contestant_number), code=CODES['error'])
        # Check if contestant is an Open Class dancer, but has not indicated that he/she has to mandatory Blind Date
        if contestant_ballroom_level == levels['open_class'] and contestant_ballroom_bd != options_yn['yes']:
            status_print('Contestant number {contestant} is an Open Class dancer in the Ballroom category, '
                         'but has NOT indicated that he/she has to Blind Date mandatory in the Ballroom category.'
                         .format(contestant=contestant_number), code=CODES['warning'])
        if contestant_latin_level == levels['open_class'] and contestant_latin_bd != options_yn['yes']:
            status_print('Contestant number {contestant} is an Open Class dancer in the Latin category, '
                         'but has NOT indicated that he/she has to Blind Date mandatory in the Latin category.'
                         .format(contestant=contestant_number), code=CODES['warning'])
        status_print('Checking contestant {num} done. Continuing check.'.format(num=contestant_number))
        status_print('')


def check_for_errors(text, code):
    """"Temp"""
    text = text.split('\n')
    if code == CODES['error']:
        text = [x for x in text if x.startswith(code)]
        if len(text) == 0:
            text = ['No errors found.']
        else:
            text.insert(0, 'Errors found:')
    if code == CODES['warning']:
        text = [x for x in text if x.startswith(code)]
        if len(text) == 0:
            text = ['No warnings found.']
        else:
            text.insert(0, 'Warnings found:')
    return text


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
            status_print('Too many team captains have signed in. The maximum is two.', code=CODES['error'])
        status_print('')
        status_print('')
        status_print('Checking each individual contestant.')
        status_print('')
        check_contestants(signup_list)
        output_text = status_text.get("1.0", END)
        status_print('')
        status_print('')
        errors = check_for_errors(output_text, code=CODES['error'])
        status_print('')
        for err in errors:
            status_print(err)
        warnings = check_for_errors(output_text, code=CODES['warning'])
        status_print('')
        status_print('')
        for war in warnings:
            status_print(war)
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
    x_scrollbar.grid(row=1, column=0, columnspan=2, padx=pad_in, sticky=E+W)
    y_scrollbar = Scrollbar(master=frame, orient=VERTICAL)
    y_scrollbar.grid(row=0, column=2, pady=pad_in, sticky=N+S)
    status_text = Text(master=frame, width=total_width, height=50, padx=pad_in, pady=pad_in,
                       xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set, state=DISABLED, wrap=NONE)
    status_text.grid(row=0, column=0, padx=pad_out, columnspan=2)
    x_scrollbar.config(command=status_text.xview)
    y_scrollbar.config(command=status_text.yview)
    pad_frame = Frame(master=frame, height=16)
    pad_frame.grid(row=2, column=0, padx=pad_in, pady=pad_out)
    start_button = Button(master=frame, text='Select file', command=select_file)
    start_button.grid(row=3, column=0, padx=pad_out, pady=pad_in)
    update_button = Button(master=frame, text='Check file', command=check_file)
    update_button.grid(row=3, column=1, padx=pad_out)
    welcome_text()
    root.mainloop()
