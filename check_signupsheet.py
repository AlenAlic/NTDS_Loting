from tkinter import *
import textwrap
from classes.entrybox import EntryBox
import os
import openpyxl
from loting import total_width, xlsx_ext, levels, options_yn, options_ymn, gen_dict
import time
import re

# TODO aangeven wanneer iemand zichzelf als partner heeft opgegeven

CODES = {'error': 'ERROR: ', 'warning': 'WARNING: '}
email_format = re.compile('^.*@.*\..*$')

file_key = {'name': '', 'path': ''}
file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['path'])
file_key['path'] = os.path.join(file_key['path'], file_key['name'])

ballroom = 'Ballroom'
latin = 'Latin'

number_not_exist = 'Contestant {con} signed up with a partner in {cat} that does not exist.'

single = 'Contestant number {con} signed up as a Blind Dater in {cat}.'
beginner_combo = 'Contestant number {con} is a Beginner in {cat}, but not in {cat2}.'

signed_up_together = 'Contestants number {con} and number {par} signed up together for {cat}.'
not_signed_up_together = 'Contestant number {con} signed up with contestant number {par} in {cat}, ' \
                         'but contestant number {par} did not sign up with contestant number {con} as well.'
opposite_roles = 'Contestants number {con} and {par} have opposite roles in {cat}.'
matching_roles = 'Contestants number {con} and {par} have selected the same role in {cat}.'
same_levels = 'Contestants number {con} and {par} are dancing at the same level in {cat}.'
different_levels = 'Contestants number {con} and {par} are dancing at different levels in {cat}.'

no_level = 'Contestant number {con} did not select a level in {cat}.'
no_role = 'Contestant number {con} did not select a role in {cat}.'
no_blind_date = 'Contestant number {con} doesn\'t have to Blind Date in {cat}, ' \
                'and is allowed to dance together with contestant number {par} in {cat}.'
blind_date = 'Contestant number {con} has to Blind Date in {cat}, ' \
             'and is not allowed to dance together with contestant number {par} in {cat}.'

no_first_name = 'Contestant number {con} has not given a first name.'
no_last_name = 'Contestant number {con} has not given a last name.'
no_email = 'Contestant number {con} has not given an e-mail address.'
email_wrong_format = 'Contestant number {con} has not given a valid e-mail address.'


def wip():
    """"Temp"""
    status_print('Work in progress...')


def column(matrix, i):
    """"Temp"""
    return [row[i] for row in matrix]


def swap(text, ch1, ch2):
    """"Temp"""
    if ch2 in ch1:
        ch1, ch2 = ch2, ch1
    text = text.replace(ch2, '!',)
    text = text.replace(ch1, ch2)
    text = text.replace('!', ch1)
    return text


def status_print(message, wrap=True, code=None):
    """"Prints the message passed to the program screen"""
    if code in list(CODES.values()):
        # if code == CODES['error']:
        message = code + message
    status_text.config(state=NORMAL)
    if wrap is True:
        message = textwrap.fill(message, total_width)
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
    status_print('To select a file you wish to check for errors, use the button below.')
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
    """"Used to select a file to check for errors."""
    ask_database = EntryBox('Enter the name of your registration file (NTDS_"TEAMNAME")', (file_key, 'name'))
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
        status_print('')
        time.sleep(0.3)
        check_file()
    elif os.path.isfile(path=file_key['path']) is False and file_key['name'] != '':
        status_print('The file "{name}" does not exist.'.format(name=file_key['name']))
        status_print('')


def check_contestants(contestants_list):
    """"Checks contestants for errors in the signup sheets."""
    for row in contestants_list:
        contestant_number = row[gen_dict['id']]
        status_print('Checking contestant {num}.'.format(num=contestant_number))
        contestant_first_name = row[gen_dict['first_name']]
        contestant_last_name = row[gen_dict['last_name']]
        contestant_email = row[gen_dict['email']]
        if contestant_first_name == '':
            status_print(no_first_name.format(con=contestant_number), code=CODES['warning'])
        if contestant_last_name == '':
            status_print(no_last_name.format(con=contestant_number), code=CODES['warning'])
        if contestant_email == '':
            status_print(no_email.format(con=contestant_number), code=CODES['warning'])
        elif email_format.match(contestant_email) is None:
                status_print(email_wrong_format.format(con=contestant_number), code=CODES['warning'])
        contestant_ballroom_level = row[gen_dict['ballroom_level']]
        contestant_ballroom_partner_number = row[gen_dict['ballroom_partner']]
        contestant_ballroom_role = row[gen_dict['ballroom_role']]
        contestant_ballroom_bd = row[gen_dict['ballroom_mandatory_blind_date']]
        contestant_latin_level = row[gen_dict['latin_level']]
        contestant_latin_partner_number = row[gen_dict['latin_partner']]
        contestant_latin_role = row[gen_dict['latin_role']]
        contestant_latin_bd = row[gen_dict['latin_mandatory_blind_date']]

        if contestant_ballroom_level == levels['beginners'] and contestant_latin_level != levels['beginners'] \
                and contestant_latin_level != '':
            status_print(beginner_combo.format(con=contestant_number, cat=ballroom, cat2=latin), code=CODES['error'])
        if contestant_ballroom_level != levels['beginners'] and contestant_ballroom_level != '' \
                and contestant_latin_level == levels['beginners']:
            status_print(beginner_combo.format(con=contestant_number, cat=latin, cat2=ballroom), code=CODES['error'])

        # Contestant signed up alone for both categories without a partner
        if contestant_ballroom_partner_number == '' and contestant_latin_partner_number == '':
            if any([all([contestant_ballroom_level != '', contestant_latin_level != '',
                         contestant_ballroom_role != '', contestant_latin_role != '']),
                    all([contestant_ballroom_level != '', contestant_latin_level == '',
                         contestant_ballroom_role != '', contestant_latin_role == '']),
                    all([contestant_ballroom_level == '', contestant_latin_level != '',
                         contestant_ballroom_role == '', contestant_latin_role != ''])]):
                if all([contestant_ballroom_level != '', contestant_latin_level != '',
                        contestant_ballroom_role != '', contestant_latin_role != '']):
                    status_print(single.format(con=contestant_number, cat=ballroom))
                    status_print(single.format(con=contestant_number, cat=latin))
                if all([contestant_ballroom_level != '', contestant_latin_level == '',
                        contestant_ballroom_role != '', contestant_latin_role == '']):
                    status_print(single.format(con=contestant_number, cat=ballroom))
                if all([contestant_ballroom_level == '', contestant_latin_level != '',
                        contestant_ballroom_role == '', contestant_latin_role != '']):
                    status_print(single.format(con=contestant_number, cat=latin))
            else:
                if contestant_ballroom_level == '' and contestant_ballroom_role != '':
                    status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                if contestant_ballroom_level != '' and contestant_ballroom_role == '':
                    status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                if contestant_latin_level == '' and contestant_latin_role != '':
                    status_print(no_role.format(con=contestant_number, cat=latin), code=CODES['error'])
                if contestant_latin_level != '' and contestant_latin_role == '':
                    status_print(no_level.format(con=contestant_number, cat=latin), code=CODES['error'])

        # Contestant signed up with a partner for both Ballroom and Latin
        elif contestant_ballroom_partner_number != '' and contestant_latin_partner_number != '':
            if contestant_ballroom_partner_number-1 > len(contestants_list):
                status_print(number_not_exist.format(con=contestant_number, cat=ballroom), code=CODES['error'])
            else:
                partner_row = contestants_list[contestant_ballroom_partner_number - 1]
                partner_number = partner_row[gen_dict['id']]
                partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
                partner_latin_number = partner_row[gen_dict['latin_partner']]
                partner_ballroom_level = partner_row[gen_dict['ballroom_level']]
                partner_latin_level = partner_row[gen_dict['latin_level']]
                partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
                partner_latin_role = partner_row[gen_dict['latin_role']]
                partner_ballroom_bd = partner_row[gen_dict['ballroom_mandatory_blind_date']]
                partner_latin_bd = partner_row[gen_dict['latin_mandatory_blind_date']]
                if contestant_latin_partner_number-1 > len(contestants_list):
                    second_partner_row = None
                else:
                    second_partner_row = contestants_list[contestant_latin_partner_number - 1]

                # Contestant signed up with the same partner for both Ballroom and Latin
                if partner_row == second_partner_row:
                    if partner_ballroom_number == contestant_number and partner_latin_number == contestant_number:
                        status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=ballroom))
                        status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=latin))
                    elif partner_ballroom_number == contestant_number and partner_latin_number != contestant_number:
                        status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=ballroom))
                        status_print(not_signed_up_together
                                     .format(con=partner_number, par=contestant_number, cat=latin), code=CODES['error'])
                    elif partner_ballroom_number != contestant_number and partner_latin_number == contestant_number:
                        status_print(not_signed_up_together
                                     .format(con=partner_number, par=contestant_number, cat=ballroom),
                                     code=CODES['error'])
                        status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=latin))
                    else:
                        status_print(not_signed_up_together
                                     .format(con=partner_number, par=contestant_number, cat=ballroom),
                                     code=CODES['error'])
                        status_print(not_signed_up_together
                                     .format(con=partner_number, par=contestant_number, cat=latin), code=CODES['error'])
                    if all([contestant_ballroom_level != '', contestant_latin_level != '',
                            partner_ballroom_level != '', partner_latin_level != '',
                            contestant_ballroom_level == partner_ballroom_level,
                            contestant_latin_level == partner_latin_level]):
                        status_print(same_levels.format(con=contestant_number, par=partner_number, cat=ballroom))
                        status_print(same_levels.format(con=contestant_number, par=partner_number, cat=latin))
                    else:
                        if contestant_ballroom_level == '':
                            status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        if contestant_latin_level == '':
                            status_print(no_level.format(con=contestant_number, cat=latin), code=CODES['error'])
                        if all([contestant_ballroom_level != '', contestant_latin_level != '',
                                partner_ballroom_level != '', partner_latin_level != '',
                                contestant_ballroom_level == partner_ballroom_level,
                                contestant_latin_level != partner_latin_level]):
                            status_print(same_levels.format(con=contestant_number, par=partner_number, cat=ballroom))
                            status_print(different_levels.format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])
                        elif all([contestant_ballroom_level != '', contestant_latin_level != '',
                                  partner_ballroom_level != '', partner_latin_level != '',
                                  contestant_ballroom_level != partner_ballroom_level,
                                  contestant_latin_level == partner_latin_level]):
                            status_print(different_levels
                                         .format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(same_levels.format(con=contestant_number, par=partner_number, cat=latin))
                        elif all([contestant_ballroom_level != '', contestant_latin_level != '',
                                  partner_ballroom_level != '', partner_latin_level != '',
                                  contestant_ballroom_level != partner_ballroom_level,
                                  contestant_latin_level != partner_latin_level]):
                            status_print(different_levels
                                         .format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(different_levels
                                         .format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])
                    if all([contestant_ballroom_role != '', contestant_latin_role != '',
                            partner_ballroom_role != '', partner_latin_role != '',
                            contestant_ballroom_role != partner_ballroom_role,
                            contestant_latin_role != partner_latin_role]):
                        status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=ballroom))
                        status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=latin))
                    else:
                        if contestant_ballroom_role == '':
                            status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        if contestant_latin_role == '':
                            status_print(no_role.format(con=contestant_number, cat=latin), code=CODES['error'])
                        if all([contestant_ballroom_role != '', contestant_latin_role != '',
                                partner_ballroom_role != '', partner_latin_role != '',
                                contestant_ballroom_role != partner_ballroom_role,
                                contestant_latin_role == partner_latin_role]):
                            status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=ballroom))
                            status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])
                        elif all([contestant_ballroom_role != '', contestant_latin_role != '',
                                  partner_ballroom_role != '', partner_latin_role != '',
                                  contestant_ballroom_role == partner_ballroom_role,
                                  contestant_latin_role != partner_latin_role]):
                            status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=latin))
                        elif all([contestant_ballroom_role != '', contestant_latin_role != '',
                                  partner_ballroom_role != '', partner_latin_role != '',
                                  contestant_ballroom_role == partner_ballroom_role,
                                  contestant_latin_role == partner_latin_role]):
                            status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])
                    if all([contestant_ballroom_bd != options_yn['yes'], contestant_latin_bd != options_yn['yes'],
                            partner_ballroom_bd != options_yn['yes'], partner_latin_bd != options_yn['yes']]):
                        status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=ballroom))
                        status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=latin))
                    else:
                        if contestant_ballroom_bd != options_yn['yes'] and contestant_latin_bd == options_yn['yes']:
                            status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=ballroom))
                            status_print(blind_date.format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])
                        elif contestant_ballroom_bd == options_yn['yes'] and contestant_latin_bd != options_yn['yes']:
                            status_print(blind_date.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=latin))
                        elif contestant_ballroom_bd == options_yn['yes'] and contestant_latin_bd == options_yn['yes']:
                            status_print(blind_date.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                            status_print(blind_date.format(con=contestant_number, par=partner_number, cat=latin),
                                         code=CODES['error'])

                # Contestant signed up with the different partners for Ballroom and Latin
                elif second_partner_row is not None:
                    second_partner_number = second_partner_row[gen_dict['id']]
                    second_partner_latin_number = second_partner_row[gen_dict['latin_partner']]
                    second_partner_latin_level = second_partner_row[gen_dict['latin_level']]
                    second_partner_latin_role = second_partner_row[gen_dict['latin_role']]
                    second_partner_latin_bd = second_partner_row[gen_dict['latin_mandatory_blind_date']]
                    if partner_ballroom_number == contestant_number:
                        status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=ballroom))
                    else:
                        status_print(not_signed_up_together
                                     .format(con=partner_number, par=contestant_number, cat=ballroom),
                                     code=CODES['error'])
                    if second_partner_latin_number == contestant_number:
                        status_print(signed_up_together
                                     .format(con=contestant_number, par=second_partner_number, cat=latin))
                    else:
                        status_print(not_signed_up_together
                                     .format(con=second_partner_number, par=contestant_number, cat=latin),
                                     code=CODES['error'])
                    if all([contestant_ballroom_level != '', partner_ballroom_level != '',
                            contestant_ballroom_level == partner_ballroom_level]):
                        status_print(same_levels.format(con=contestant_number, par=partner_number, cat=ballroom))
                    else:
                        if contestant_ballroom_level == '':
                            status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        elif contestant_ballroom_level != partner_ballroom_level and partner_ballroom_level != '':
                            status_print(
                                different_levels.format(con=contestant_number, par=partner_number, cat=ballroom),
                                code=CODES['error'])
                    if all([contestant_latin_level != '', second_partner_latin_level != '',
                            contestant_latin_level == second_partner_latin_level]):
                        status_print(
                            same_levels.format(con=contestant_number, par=second_partner_number, cat=latin))
                    else:
                        if contestant_latin_level == '':
                            status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        elif contestant_latin_level != second_partner_latin_level and second_partner_latin_level != '':
                            status_print(
                                different_levels.format(con=contestant_number, par=second_partner_number, cat=latin),
                                code=CODES['error'])
                    if all([contestant_ballroom_role != '', partner_ballroom_role != '',
                            contestant_ballroom_role != partner_ballroom_role]):
                        status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=ballroom))
                    else:
                        if contestant_ballroom_role == '':
                            status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        elif contestant_ballroom_role != partner_ballroom_role and partner_ballroom_role != '':
                            status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                    if all([contestant_latin_role != '', second_partner_latin_role != '',
                            contestant_latin_role != second_partner_latin_role]):
                        status_print(opposite_roles.format(con=contestant_number, par=second_partner_number, cat=latin))
                    else:
                        if contestant_latin_role == '':
                            status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                        elif contestant_latin_role != second_partner_latin_role and second_partner_latin_role != '':
                            status_print(matching_roles
                                         .format(con=contestant_number, par=second_partner_number, cat=latin),
                                         code=CODES['error'])
                    if contestant_ballroom_bd != options_yn['yes'] and partner_ballroom_bd != options_yn['yes']:
                        status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=ballroom))
                    elif contestant_ballroom_bd == options_yn['yes']:
                            status_print(blind_date.format(con=contestant_number, par=partner_number, cat=ballroom),
                                         code=CODES['error'])
                    if contestant_latin_bd != options_yn['yes'] and second_partner_latin_bd != options_yn['yes']:
                        status_print(no_blind_date.format(con=contestant_number, par=second_partner_number, cat=latin))
                    elif contestant_latin_bd == options_yn['yes']:
                            status_print(blind_date.format(con=contestant_number, par=second_partner_number, cat=latin),
                                         code=CODES['error'])
                else:
                    status_print(number_not_exist.format(con=contestant_number, cat=latin), code=CODES['error'])

        # Contestant signed up dancing only Ballroom with a partner
        elif contestant_ballroom_partner_number != '' and contestant_latin_partner_number == '':
            partner_row = contestants_list[contestant_ballroom_partner_number - 1]
            partner_number = partner_row[gen_dict['id']]
            partner_ballroom_number = partner_row[gen_dict['ballroom_partner']]
            partner_ballroom_level = partner_row[gen_dict['ballroom_level']]
            partner_ballroom_role = partner_row[gen_dict['ballroom_role']]
            partner_ballroom_bd = partner_row[gen_dict['ballroom_mandatory_blind_date']]
            if partner_ballroom_number == contestant_number:
                status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=ballroom))
            else:
                status_print(not_signed_up_together.format(con=contestant_number, par=partner_number, cat=ballroom),
                             code=CODES['error'])
            if all([contestant_ballroom_level != '', partner_ballroom_level != '',
                    contestant_ballroom_level == partner_ballroom_level]):
                status_print(same_levels.format(con=contestant_number, par=partner_number, cat=ballroom))
            else:
                if contestant_ballroom_level == '':
                    status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                elif contestant_ballroom_level != partner_ballroom_level and partner_ballroom_level != '':
                    status_print(
                        different_levels.format(con=contestant_number, par=partner_number, cat=ballroom),
                        code=CODES['error'])
            if all([contestant_ballroom_role != '', partner_ballroom_role != '',
                    contestant_ballroom_role != partner_ballroom_role]):
                status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=ballroom))
            else:
                if contestant_ballroom_role == '':
                    status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                elif contestant_ballroom_role != partner_ballroom_role:
                    status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=ballroom),
                                 code=CODES['error'])
            if contestant_ballroom_bd != options_yn['yes'] and partner_ballroom_bd != options_yn['yes']:
                status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=ballroom))
            elif contestant_ballroom_bd == options_yn['yes']:
                status_print(blind_date.format(con=contestant_number, par=partner_number, cat=ballroom),
                             code=CODES['error'])
            if contestant_latin_level != '' and contestant_latin_role != '':
                status_print(single.format(con=contestant_number, cat=latin))
            elif contestant_latin_level != '' and contestant_latin_role == '':
                status_print(no_role.format(con=contestant_number, cat=latin), code=CODES['error'])
            elif contestant_latin_level == '' and contestant_latin_role != '':
                status_print(no_level.format(con=contestant_number, cat=latin), code=CODES['error'])

        # Contestant signed up dancing only Latin with a partner
        elif contestant_ballroom_partner_number == '' and contestant_latin_partner_number != '':
            partner_row = contestants_list[contestant_latin_partner_number - 1]
            partner_number = partner_row[gen_dict['id']]
            partner_latin_number = partner_row[gen_dict['latin_partner']]
            partner_latin_level = partner_row[gen_dict['latin_level']]
            partner_latin_role = partner_row[gen_dict['latin_role']]
            partner_latin_bd = partner_row[gen_dict['latin_mandatory_blind_date']]
            if contestant_ballroom_level != '' and contestant_ballroom_role != '':
                status_print(single.format(con=contestant_number, cat=ballroom))
            elif contestant_ballroom_level != '' and contestant_ballroom_role == '':
                status_print(no_role.format(con=contestant_number, cat=ballroom), code=CODES['error'])
            elif contestant_ballroom_level == '' and contestant_ballroom_role != '':
                status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
            if partner_latin_number == contestant_number:
                status_print(signed_up_together.format(con=contestant_number, par=partner_number, cat=latin))
            else:
                status_print(not_signed_up_together.format(con=contestant_number, par=partner_number, cat=latin),
                             code=CODES['error'])
            if all([contestant_latin_level != '', partner_latin_level != '',
                    contestant_latin_level == partner_latin_level]):
                status_print(same_levels.format(con=contestant_number, par=partner_number, cat=latin))
            else:
                if contestant_latin_level == '':
                    status_print(no_level.format(con=contestant_number, cat=ballroom), code=CODES['error'])
                elif contestant_latin_level != partner_latin_level and partner_latin_level != '':
                    status_print(
                        different_levels.format(con=contestant_number, par=partner_number, cat=latin),
                        code=CODES['error'])
            if all([contestant_latin_role != '', partner_latin_role != '',
                    contestant_latin_role != partner_latin_role]):
                status_print(opposite_roles.format(con=contestant_number, par=partner_number, cat=latin))
            else:
                if contestant_latin_role == '':
                    status_print(no_role.format(con=contestant_number, cat=latin), code=CODES['error'])
                elif contestant_latin_role != partner_latin_role and partner_latin_role != '':
                    status_print(matching_roles.format(con=contestant_number, par=partner_number, cat=latin),
                                 code=CODES['error'])
            if contestant_latin_bd != options_yn['yes'] and partner_latin_bd != options_yn['yes']:
                status_print(no_blind_date.format(con=contestant_number, par=partner_number, cat=latin))
            elif contestant_latin_bd == options_yn['yes']:
                status_print(blind_date.format(con=contestant_number, par=partner_number, cat=latin),
                             code=CODES['error'])

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
    """"Returns list of errors."""
    text = text.split('\n')
    if code == CODES['error']:
        text = [x for x in text if x.startswith(code)]
        text = check_duplicates(text)
        if len(text) == 0:
            text = ['No errors found.']
        else:
            text.insert(0, 'ERRORS FOUND:')
    if code == CODES['warning']:
        text = [x for x in text if x.startswith(code)]
        text = check_duplicates(text)
        if len(text) == 0:
            text = ['No warnings found.']
        else:
            text.insert(0, 'WARNINGS FOUND:')
    return text


def check_duplicates(text):
    """Checks for duplicates in the error text (ex: error with 2 and 6, and 6 and 2) and returns a list without them."""
    non_duplicates = list()
    for t in text:
        warning_numbers = [int(s) for s in t.split() if s.isdigit()]
        if len(warning_numbers) > 2:
            warning_numbers = list(set(warning_numbers))
        if len(warning_numbers) == 2:
            test_str = swap(t, str(warning_numbers[1]), str(warning_numbers[0]))
            if test_str not in non_duplicates:
                non_duplicates.append(t)
        elif len(warning_numbers) == 1:
            if t not in non_duplicates:
                non_duplicates.append(t)
    return non_duplicates


def check_file():
    """Main program"""
    update_button.config(state=DISABLED)
    status_print('Checking file: {file}'.format(file=file_key['path']))
    time.sleep(1)
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
    status_print('')
    status_print('')
    if len(errors) == 1 and len(warnings) == 1:
        status_print('No errors or warnings were found in the file "{file}".'.format(file=file_key['path']))
        status_print('The tested signup sheet is OK.')
    else:
        status_print('One or more errors and/or warnings were found. Please see the list printed above.')
    update_button.config(state=NORMAL)


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
    status_text = Text(master=frame, width=total_width, height=50, padx=pad_in, pady=pad_in,
                       xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set, state=DISABLED, wrap=NONE)
    status_text.grid(row=0, column=0, padx=pad_out)
    x_scrollbar.config(command=status_text.xview)
    y_scrollbar.config(command=status_text.yview)
    pad_frame = Frame(master=frame, height=16)
    pad_frame.grid(row=2, column=0, padx=pad_in, pady=pad_out)
    update_button = Button(master=frame, text='Check file', command=select_file)
    update_button.grid(row=3, column=0, padx=pad_out)
    welcome_text()
    root.mainloop()
