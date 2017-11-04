# Main file for the program
import sqlite3 as sql
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from random import randint
import random
import time
from tabulate import tabulate
from tkinter import *
from tkinter import messagebox
from classes.entrybox import EntryBox
import os.path
import configparser
import textwrap
import shutil

# Program constants
total_width = 190
status_text_width = 140
default_db_name = 'NTDS'
db_ext = '.db'
xlsx_ext = '.xlsx'
database_key = {'db': '', 'path': ''}

# Setup configuration and setting files
config_folder = 'config'
settings_folder = 'settings'
statistics_folder = 'statistics'
config_parser = configparser.ConfigParser()
config_key = {'name': 'config.ini', 'folder': config_folder, 'path': ''}
config_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_key['folder'])
config_key['path'] = os.path.join(config_key['path'], config_key['name'])
settings_key = {'name': 'user_settings.ini', 'folder': settings_folder, 'path': ''}
settings_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), settings_key['folder'])
settings_key['path'] = os.path.join(settings_key['path'], settings_key['name'])
participating_teams_key = {'name': 'NTDS_participating_teams.ini', 'folder': config_folder, 'path': ''}
participating_teams_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                               participating_teams_key['folder'])
participating_teams_key['path'] = os.path.join(participating_teams_key['path'], participating_teams_key['name'])
template_key = {'name': 'NTDS_Template.xlsx', 'folder': config_folder, 'path': ''}
template_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), template_key['folder'])
template_key['path'] = os.path.join(template_key['path'], template_key['name'])
output_key = {'name': 'NTDS_Selection', 'folder': 'output', 'path': ''}
output_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_key['folder'])
output_key['path'] = os.path.join(output_key['path'], output_key['name'])

# NODIG
# TODO controle programma voor inschrijflijst
# TODO add additional teams without crashing the system
# TODO dummy
# WISHLIST
# TODO export, add formatting
# TODO refactor (remove hardcoded material that snuck in)


# Dit kan onder elk kopje gebeur
# TODO create files and folders on first run
# TODO config and user settings
# Create files if they are not there
if not os.path.isfile(path=settings_key['path']):
    open(settings_key['path'], 'w+').close()
if not os.path.isfile(path=config_key['path']):
    open(config_key['path'], 'w+').close()

# Boundaries
user_boundaries = False
max_contestants = 400
max_guaranteed_beginners = 4
max_fixed_lion_contestants = 10
beginner_signup_cutoff = int(max_contestants*20/100)
buffer_for_selection = 40
if os.path.isfile(path=settings_key['path']):
    config_parser.read(settings_key['path'])
    if 'ContestantNumbers' in config_parser:
        user_boundaries = True
        max_contestants = config_parser['ContestantNumbers'].getint('MaximumNumberOfContestants')
        max_guaranteed_beginners = config_parser['ContestantNumbers'].getint('MinimumGuaranteedBeginners')
        max_fixed_lion_contestants = config_parser['ContestantNumbers'].getint('MinimumGuaranteedLions')
        beginner_signup_cutoff = int(max_contestants * config_parser['ContestantNumbers']
                                     .getint('BeginnerCutoffPercentage') / 100)
        buffer_for_selection = config_parser['ContestantNumbers'].getint('BufferForManualSelection')
    else:
        print('Dummy')
        # Update file

# Levels
levels = {'beg': 'Beginners', 'breiten': 'Breitensport', 'closed': 'CloseD', 'open': 'Open Class', 0: ''}
if os.path.isfile(path=config_key['path']):
    config_parser.read(config_key['path'])
    if 'DancingClasses' in config_parser:
        levels = dict(config_parser.items('DancingClasses'))
    else:
        print('Dummy')
        # Update file

# Participating classes
classes = [levels['beg'], levels['breiten'], levels['open']]
if os.path.isfile(path=settings_key['path']):
    if 'AvailableClasses' in config_parser:
        classes = list()
        for dancing_class in config_parser['AvailableClasses']:
            if config_parser['AvailableClasses'].getboolean(dancing_class):
                classes.append(levels[dancing_class])
    else:
        print('Dummy')
        # Update file

# Roles
roles = {'lead': 'Lead', 'follow': 'Follow', 0: ''}
if os.path.isfile(path=config_key['path']):
    config_parser.read(config_key['path'])
    if 'DancingClasses' in config_parser:
        roles = dict(config_parser.items('DancingRoles'))
    else:
        print('Dummy')
        # Update file

# Signup options
yes = 'Ja'
maybe = 'Misschien'
no = 'Nee'
options_ymn = options_ym = {'yes': yes, 'maybe': maybe, 'no': no}
option_yn = {'yes': yes, 'no': no}
if os.path.isfile(path=settings_key['path']):
    if 'YesMaybeNo' in config_parser:
        yes = config_parser['YesMaybeNo']['Yes']
        maybe = config_parser['YesMaybeNo']['Maybe']
        no = config_parser['YesMaybeNo']['No']
        options_ymn = dict(config_parser.items('YesMaybeNo'))
        option_yn = {'yes': yes, 'no': no}
    else:
        print('Dummy')
        # Update file

# Participants Lion points
lion_participants = [levels['beg'], levels['breiten']]
if os.path.isfile(path=settings_key['path']):
    if 'Lions' in config_parser:
        lion_participants = list()
        for lion_level in config_parser['Lions']:
            if config_parser['Lions'].getboolean(lion_level):
                lion_participants.append(levels[lion_level])
    else:
        print('Dummy')
        # Update file

# Statistics options
gather_stats = False
runs = 1000
if os.path.isfile(path=settings_key['path']):
    if 'SelectionMode' in config_parser:
        gather_stats = config_parser['SelectionMode'].getboolean('StatisticalAnalysis', gather_stats)
        runs = config_parser['SelectionMode'].getint('NumberOfRunsForStatisticalAnalysis', runs)
    else:
        print('Dummy')
        # Update file

# Title for Excel export files
output_title = []
title_book = openpyxl.load_workbook(template_key['path'])
title_sheet = title_book.worksheets[0]
title_col = 0
while True:
    title_col += 1
    if title_sheet.cell(row=1, column=title_col).value is None:
        break
    else:
        output_title.append(title_sheet.cell(row=1, column=title_col).value)

# Table names
signup_list = 'signup_list'
cancelled_list = 'cancelled_list'
selection_list = 'selection_list'
selected_list = 'selected_list'
backup_list = 'backup_list'
team_list = 'team_list'
partners_list = 'partners_list'
ref_partner_list = 'reference_partner_list'
fixed_beginners_list = 'fixed_beginners'
fixed_lions_list = 'fixed_lions'
beginners_list = 'beginners'
lions_list = 'lions'
contestants_list = 'contestants'
individual_list = 'individuals'

# Formatting names
lions = 'Lions'
contestants = 'contestants'
individuals = 'individuals'

# Selectable options
NTDS_ballroom_level = levels
NTDS_latin_level = levels
NTDS_ballroom_role = roles
NTDS_latin_role = roles
NTDS_ballroom_mandatory_blind_date = option_yn
NTDS_latin_mandatory_blind_date = option_yn
NTDS_team_captain = option_yn
NTDS_first_aid = options_ymn
NTDS_emergency_response_officer = options_ymn
NTDS_ballroom_jury = options_ymn
NTDS_latin_jury = options_ymn
NTDS_student = option_yn
NTDS_sleeping_location = {'yes': yes, 0: ''}
NTDS_current_volunteer = options_ymn
NTDS_past_volunteer = option_yn
NTDS_options = {'ballroom_level': NTDS_ballroom_level, 'latin_level': NTDS_latin_level,
                'ballroom_role': NTDS_ballroom_role, 'latin_role': NTDS_latin_role,
                'ballroom_mandatory_blind_date': NTDS_ballroom_mandatory_blind_date,
                'latin_mandatory_blind_date': NTDS_latin_mandatory_blind_date,
                'team_captain': NTDS_team_captain,
                'first_aid': NTDS_first_aid, 'emergency_response_officer': NTDS_emergency_response_officer,
                'ballroom_jury': NTDS_ballroom_jury, 'latin_jury': NTDS_latin_jury,
                'student': NTDS_student, 'sleeping_location': NTDS_sleeping_location,
                'current_volunteer': NTDS_current_volunteer, 'past_volunteer': NTDS_past_volunteer}

# SQL Table column names and dictionary for dancers lists
sql_id = 'id'
sql_first_name = 'first_name'
sql_ln_prefix = 'ln_prefix'
sql_last_name = 'last_name'
sql_email = 'email'
sql_ballroom_level = 'ballroom_level'
sql_latin_level = 'latin_level'
sql_ballroom_partner = 'ballroom_partner'
sql_latin_partner = 'latin_partner'
sql_role = 'role'
sql_ballroom_role = 'ballroom_role'
sql_latin_role = 'latin_role'
sql_team_captain = 'team_captain'
sql_ballroom_mandatory_blind_date = 'ballroom_mandatory_blind_date'
sql_latin_mandatory_blind_date = 'latin_mandatory_blind_date'
sql_first_aid = 'first_aid'
sql_emergency_response_officer = 'emergency_response_officer'
sql_ballroom_jury = 'ballroom_jury'
sql_latin_jury = 'latin_jury'
sql_student = 'student'
sql_sleeping_location = 'sleeping_location'
sql_diet = 'diet_wishes'
sql_current_volunteer = 'current_volunteer'
sql_previous_volunteer = 'past_volunteer'
sql_city = 'city'
gen_dict = {sql_id: 0, sql_first_name: 1, sql_ln_prefix: 2, sql_last_name: 3, sql_email: 4,
            sql_ballroom_level: 5, sql_latin_level: 6, sql_ballroom_partner: 7, sql_latin_partner: 8,
            sql_ballroom_role: 9, sql_latin_role: 10,
            sql_ballroom_mandatory_blind_date: 11, sql_latin_mandatory_blind_date: 12,  sql_team_captain: 13,
            sql_first_aid: 14, sql_emergency_response_officer: 15, sql_ballroom_jury: 16, sql_latin_jury: 17,
            sql_student: 18, sql_sleeping_location: 19, sql_diet: 20,
            sql_current_volunteer: 21, sql_previous_volunteer: 22,
            sql_city: 23}

# SQL Table column names and dictionary of teams list
sql_team = 'team'
sql_signup_list = 'signup_list'
team_dict = {sql_team: 0, sql_city: 1, sql_signup_list: 2}

# SQL Table column names and dictionary for partners list
sql_no = 'num'
sql_lead = 'lead'
sql_follow = 'follow'
sql_city_lead = 'city_lead'
sql_city_follow = 'city_follow'
sql_ballroom_level_lead = 'ballroom_level_lead'
sql_ballroom_level_follow = 'ballroom_level_follow'
sql_latin_level_lead = 'latin_level_lead'
sql_latin_level_follow = 'latin_level_follow'
partner_dict = {sql_no: 0, sql_lead: 1, sql_follow: 2, sql_city_lead: 3, sql_city_follow: 4,
                sql_ballroom_level_lead: 5, sql_ballroom_level_follow: 6,
                sql_latin_level_lead: 7, sql_latin_level_follow: 8}

# SQL Table column names and dictionary for city list
sql_num_con = 'number_of_contestants'
sql_max_con = 'max_contestants'
city_dict = {sql_city: 0, sql_num_con: 1, sql_max_con: 2}

# SQL Table column names for individuals list
sql_run = 'run_number'

# TODO dit automatiseren?
# General query formats
drop_table_query = 'DROP TABLE IF EXISTS {};'
dancers_list_query = 'CREATE TABLE {table_name} ({id} INT PRIMARY KEY, ' \
                     '{first_name} TEXT, {ln_prefix} TEXT, {last_name} TEXT, {email} TEXT, ' \
                     '{ballroom_level} TEXT, {latin_level} TEXT, {ballroom_partner} INT, {latin_partner} INT, ' \
                     '{ballroom_role} TEXT, {latin_role} TEXT, ' \
                     '{ballroom_mandatory_blind_date} TEXT, {latin_mandatory_blind_date} TEXT, {team_captain} TEXT,' \
                     '{first_aid} TEXT, {emergency_response_officer} TEXT, {ballroom_jury} TEXT, {latin_jury} TEXT, ' \
                     '{student} TEXT, {sleeping_location} TEXT, {diet} TEXT, ' \
                     '{current_volunteer} TEXT, {previous_volunteer} TEXT, ' \
                     '{city} TEXT);'\
     .format(table_name={}, id=sql_id, first_name=sql_first_name, ln_prefix=sql_ln_prefix, last_name=sql_last_name,
             email=sql_email, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
             ballroom_partner=sql_ballroom_partner, latin_partner=sql_latin_partner,
             ballroom_role=sql_ballroom_role, latin_role=sql_latin_role, team_captain=sql_team_captain,
             ballroom_mandatory_blind_date=sql_ballroom_mandatory_blind_date,
             latin_mandatory_blind_date=sql_latin_mandatory_blind_date,
             first_aid=sql_first_aid, emergency_response_officer=sql_emergency_response_officer,
             ballroom_jury=sql_ballroom_jury, latin_jury=sql_latin_jury, student=sql_student,
             sleeping_location=sql_sleeping_location, diet=sql_diet,
             current_volunteer=sql_current_volunteer, previous_volunteer=sql_previous_volunteer,
             city=sql_city)
team_list_query = 'CREATE TABLE {tn} ({team} TEXT PRIMARY KEY, {city} TEXT, {signup_list} TEXT);' \
    .format(tn={}, team=sql_team, city=sql_city, signup_list=sql_signup_list)
paren_table_query = 'CREATE TABLE {tn} ({no} INTEGER PRIMARY KEY AUTOINCREMENT, ' \
                    '{lead} INT, {follow} INT, {lead_city} TEXT, {follow_city} TEXT, ' \
                    '{ballroom_level_lead} TEXT, {ballroom_level_follow} TEXT, ' \
                    '{latin_level_lead} TEXT, {latin_level_follow} TEXT);'\
    .format(tn={}, no=sql_no, lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow,
            ballroom_level_lead=sql_ballroom_level_lead, ballroom_level_follow=sql_ballroom_level_follow,
            latin_level_lead=sql_latin_level_lead, latin_level_follow=sql_latin_level_follow)
city_list_query = 'CREATE TABLE {tn} ({city} TEXT, {num} INT, {max_num} INT);' \
    .format(tn={}, city=sql_city, num=sql_num_con, max_num=sql_max_con)


def find_partner(identifier, connection, cursor, city=None, signed_partner_only=False):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    status_print('Looking for a partner for dancer {id}'.format(id=identifier))
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
    dancer = cursor.execute(query, (identifier,)).fetchone()
    if dancer is not None:
        team = dancer[gen_dict[sql_city]]
        ballroom_level = dancer[gen_dict[sql_ballroom_level]]
        latin_level = dancer[gen_dict[sql_latin_level]]
        if ballroom_level == '':
            ballroom_role = ''
        else:
            ballroom_role = dancer[gen_dict[sql_ballroom_role]]
        if latin_level == '':
            latin_role = ''
        else:
            latin_role = dancer[gen_dict[sql_latin_role]]
        ballroom_partner = dancer[gen_dict[sql_ballroom_partner]]
        latin_partner = dancer[gen_dict[sql_latin_partner]]
        # Check if the contestant already has signed up with a partner (or two)
        if isinstance(ballroom_partner, int) and ballroom_partner == latin_partner:
            partner_id = ballroom_partner
        if isinstance(ballroom_partner, int) and latin_partner == '':
            partner_id = ballroom_partner
            query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
            partner = cursor.execute(query, (partner_id,)).fetchone()
            partner_latin_partner = partner[gen_dict[sql_latin_partner]]
            if isinstance(partner_latin_partner, int):
                status_print('{id1} and {id2} signed up together'
                             .format(id1=ballroom_partner, id2=partner_latin_partner))
                create_pair(ballroom_partner, partner_latin_partner, connection=connection, cursor=cursor)
                move_selected_contestant(partner_latin_partner, connection=connection, cursor=cursor)
        if ballroom_partner == '' and isinstance(latin_partner, int):
            partner_id = latin_partner
            query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
            partner = cursor.execute(query, (partner_id,)).fetchone()
            partner_ballroom_partner = partner[gen_dict[sql_ballroom_partner]]
            if isinstance(partner_ballroom_partner, int):
                status_print('{id1} and {id2} signed up together'
                             .format(id1=latin_partner, id2=partner_ballroom_partner))
                create_pair(latin_partner, partner_ballroom_partner, connection=connection, cursor=cursor)
                move_selected_contestant(partner_ballroom_partner, connection=connection, cursor=cursor)
        if all([isinstance(ballroom_partner, int), isinstance(latin_partner, int), ballroom_partner != latin_partner]):
            status_print('{id1} and {id2} signed up together'.format(id1=identifier, id2=ballroom_partner))
            create_pair(identifier, ballroom_partner, connection=connection, cursor=cursor)
            move_selected_contestant(ballroom_partner, connection=connection, cursor=cursor)
            partner_id = latin_partner
        if partner_id is not None:
            status_print('{id1} and {id2} signed up together'.format(id1=identifier, id2=partner_id))
        if partner_id is None:
            query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {latin_level} = ? AND {ballroom_partner} = "" ' \
                    'AND {latin_partner} = "" ' \
                .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                        ballroom_partner=sql_ballroom_partner, latin_partner=sql_latin_partner)
            if ballroom_role != '' and ballroom_role == latin_role:
                query += 'AND {ballroom_role} != ? AND latin_role != ? '\
                    .format(ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
            elif ballroom_role != '' and latin_role == '':
                query += 'AND {ballroom_role} != ? AND latin_role = ? ' \
                    .format(ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
            elif ballroom_role == '' and latin_role != '':
                query += 'AND {ballroom_role} = ? AND latin_role != ? ' \
                    .format(ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
            elif ballroom_role == '' and latin_role == '':
                query += 'AND {ballroom_role} = ? AND latin_role = ? ' \
                    .format(ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
            else:
                query += 'AND {ballroom_role} = ? AND latin_role != ? ' \
                    .format(ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
                query2 = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {latin_level} = ? AND ' \
                         '{ballroom_partner} = "" AND {latin_partner} = "" ' \
                         'AND {ballroom_role} = ? AND latin_role = ? '\
                    .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                            ballroom_partner=sql_ballroom_partner, latin_partner=sql_latin_partner,
                            ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
                if city is None:
                    query2 += ' AND {team} != ?'.format(team=sql_city)
                else:
                    query2 += ' AND {team} = ?'.format(team=sql_city)
                    team = city
                potential_partners = cursor.execute(query2, (ballroom_level, '', ballroom_role, '', team)).fetchall()
                number_of_potential_partners = len(potential_partners)
                if number_of_potential_partners > 0:
                    random_num = randint(0, number_of_potential_partners - 1)
                    first_partner_id = potential_partners[random_num][gen_dict[sql_id]]
                    if first_partner_id is not None:
                        status_print('Different roles: {id1} and {id2} matched together'
                                     .format(id1=identifier, id2=first_partner_id))
                        create_pair(identifier, first_partner_id, connection=connection, cursor=cursor)
                        move_selected_contestant(first_partner_id, connection=connection, cursor=cursor)
                ballroom_level = ''
                ballroom_role = ''
            if city is None:
                query += ' AND {team} != ?'.format(team=sql_city)
            else:
                query += ' AND {team} = ?'.format(team=sql_city)
                team = city
        # Try to find an "ideal" partner with the same combination of levels for the dancer
        potential_partners = []
        number_of_potential_partners = len(potential_partners)
        if signed_partner_only is False:
            if partner_id is None:
                potential_partners = cursor.\
                    execute(query, (ballroom_level, latin_level, ballroom_role, latin_role, team))\
                    .fetchall()
                number_of_potential_partners = len(potential_partners)
                if number_of_potential_partners > 0:
                    random_num = randint(0, number_of_potential_partners - 1)
                    partner_id = potential_partners[random_num][gen_dict[sql_id]]
            query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {latin_level} = ? AND {ballroom_partner} = "" ' \
                    'AND {latin_partner} = "" AND {ballroom_role} != ? AND latin_role != ? '\
                .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                        ballroom_partner=sql_ballroom_partner, latin_partner=sql_latin_partner,
                        ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
            if city is None:
                query += ' AND {team} != ?'.format(team=sql_city)
            else:
                query += ' AND {team} = ?'.format(team=sql_city)
                team = city
            # Try to find a partner for a beginner, beginner combination
            if all([ballroom_level == levels['beg'], latin_level == levels['beg'],
                    partner_id is None, number_of_potential_partners == 0]):
                potential_partners += cursor.execute(query, (levels['beg'], '', ballroom_role, '', team)).fetchall()
                potential_partners += cursor.execute(query, ('', levels['beg'], '', latin_role, team)).fetchall()
            # Try to find a partner for a levels['beg'], Null combination
            if all([ballroom_level == levels['beg'], latin_level == '',
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.\
                    execute(query, (levels['beg'], levels['beg'], ballroom_role, ballroom_role, team)).fetchall()
            # Try to find a partner for a Null, beginner combination
            if all([ballroom_level == '', latin_level == levels['beg'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.\
                    execute(query, (levels['beg'], levels['beg'], latin_role, latin_role, team)).fetchall()
            # Try to find a partner for a breiten, breiten combination
            if all([ballroom_level == levels['breiten'], latin_level == levels['breiten'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], '', ballroom_role, '', team)).fetchall()
                potential_partners += cursor.execute(query, ('', levels['breiten'], '', latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['breiten'], levels['open'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['breiten'],
                                                             ballroom_role, latin_role, team)).fetchall()
            # Try to find a partner for a breiten, Null combination
            if all([ballroom_level == levels['breiten'], latin_level == '',
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], levels['breiten'],
                                                             ballroom_role, ballroom_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['breiten'], levels['open'],
                                                             ballroom_role, ballroom_role, team)).fetchall()
            # Try to find a partner for a Null, Breiten combination
            if all([ballroom_level == '', latin_level == levels['breiten'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], levels['breiten'],
                                                             latin_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['breiten'],
                                                             latin_role, latin_role, team)).fetchall()
            # Try to find a partner for a Breiten, Open combination
            if all([ballroom_level == levels['breiten'], latin_level == levels['open'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], levels['breiten'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['breiten'], '', ballroom_role, '', team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['open'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, ('', levels['open'], '', latin_role, team)).fetchall()
            # Try to find a partner for a Open, Breiten combination
            if all([ballroom_level == levels['open'], latin_level == levels['breiten'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], levels['breiten'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, ('', levels['breiten'], '', latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['open'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], '', ballroom_role, '', team)).fetchall()
            # Try to find a partner for a Open, Open combination
            if all([ballroom_level == levels['open'], latin_level == levels['open'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['breiten'], levels['open'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['breiten'],
                                                             ballroom_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], '', ballroom_role, '', team)).fetchall()
                potential_partners += cursor.execute(query, ('', levels['open'], '', latin_role, team)).fetchall()
            # Try to find a partner for a Open, Null combination
            if all([ballroom_level == levels['open'], latin_level == '',
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor.execute(query, (levels['open'], levels['breiten'],
                                                             ballroom_role, ballroom_role, team)).fetchall()
                potential_partners += cursor\
                    .execute(query, (levels['open'], levels['open'], ballroom_role, ballroom_role, team)).fetchall()
            # Try to find a partner for a Null, Open combination
            if all([ballroom_level == '', latin_level == levels['open'],
                    number_of_potential_partners == 0, partner_id is None]):
                potential_partners += cursor\
                    .execute(query, (levels['breiten'], levels['open'], latin_role, latin_role, team)).fetchall()
                potential_partners += cursor.execute(query, (levels['open'], levels['open'],
                                                             latin_role, latin_role, team)).fetchall()
        # If there is a potential partner, randomly select one
        if partner_id is None:
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        if partner_id is None:
            status_print('Found no match for {id1}'.format(id1=identifier))
        elif partner_id is not None and ballroom_partner == '' and latin_partner == '':
            status_print('Matched {id1} and {id2} together'.format(id1=identifier, id2=partner_id))
    return partner_id


def create_pair(first_dancer, second_dancer, connection, cursor):
    """Creates a pair of the two selected dancers and writes their data away in partner lists"""
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=signup_list, id=sql_id)
    first_dancer = cursor.execute(query, (first_dancer,)).fetchone()
    second_dancer = cursor.execute(query, (second_dancer,)).fetchone()
    if first_dancer is None:
        first_dancer_ballroom_role = ''
        first_dancer_latin_role = ''
    else:
        first_dancer_ballroom_role = first_dancer[gen_dict[sql_ballroom_role]]
        first_dancer_latin_role = first_dancer[gen_dict[sql_latin_role]]
    if first_dancer_ballroom_role == roles['follow'] or first_dancer_latin_role == roles['follow']:
        first_dancer, second_dancer = second_dancer, first_dancer
    if first_dancer is None:
        first_dancer_id = ''
        first_dancer_team = ''
        first_dancer_ballroom_level = ''
        first_dancer_latin_level = ''
    else:
        first_dancer_id = first_dancer[gen_dict[sql_id]]
        first_dancer_team = first_dancer[gen_dict[sql_city]]
        first_dancer_ballroom_level = first_dancer[gen_dict[sql_ballroom_level]]
        first_dancer_latin_level = first_dancer[gen_dict[sql_latin_level]]
    if second_dancer is None:
        second_dancer_id = ''
        second_dancer_team = ''
        second_dancer_ballroom_level = ''
        second_dancer_latin_level = ''
    else:
        second_dancer_id = second_dancer[gen_dict[sql_id]]
        second_dancer_team = second_dancer[gen_dict[sql_city]]
        second_dancer_ballroom_level = second_dancer[gen_dict[sql_ballroom_level]]
        second_dancer_latin_level = second_dancer[gen_dict[sql_latin_level]]
    query = 'INSERT INTO {tn} ({lead}, {follow}, {lead_city}, {follow_city}, ' \
            '{ballroom_level_lead}, {ballroom_level_follow}, {latin_level_lead}, {latin_level_follow}) ' \
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?)'\
        .format(tn=partners_list,
                lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow,
                ballroom_level_lead=sql_ballroom_level_lead, ballroom_level_follow=sql_ballroom_level_follow,
                latin_level_lead=sql_latin_level_lead, latin_level_follow=sql_latin_level_follow)
    cursor.execute(query, (first_dancer_id, second_dancer_id, first_dancer_team, second_dancer_team,
                           first_dancer_ballroom_level, second_dancer_ballroom_level,
                           first_dancer_latin_level, second_dancer_latin_level))
    query = 'INSERT INTO {tn} ({lead}, {follow}, {lead_city}, {follow_city}, ' \
            '{ballroom_level_lead}, {ballroom_level_follow}, {latin_level_lead}, {latin_level_follow}) ' \
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?)' \
        .format(tn=ref_partner_list,
                lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow,
                ballroom_level_lead=sql_ballroom_level_lead, ballroom_level_follow=sql_ballroom_level_follow,
                latin_level_lead=sql_latin_level_lead, latin_level_follow=sql_latin_level_follow)
    cursor.execute(query, (first_dancer_id, second_dancer_id, first_dancer_team, second_dancer_team,
                           first_dancer_ballroom_level, second_dancer_ballroom_level,
                           first_dancer_latin_level, second_dancer_latin_level))
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=selection_list, id=sql_id)
    cursor.executemany(query, [(first_dancer_id,), (second_dancer_id,)])
    connection.commit()


def move_selected_contestant(identifier, connection, cursor):
    """Moves dancer, given id, from the selection list to selected list"""
    if identifier is not None:
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE id = ?;'.format(tn1=selected_list, tn2=signup_list)
        cursor.execute(query, (identifier,))
        query = 'DELETE FROM {tn} WHERE id = ?'.format(tn=selection_list)
        cursor.execute(query, (identifier,))
        connection.commit()
        status_print('Selected {} for the NTDS'.format(identifier))


def remove_selected_contestant(identifier, connection, cursor):
    """Temp"""
    if identifier is not None:
        query = 'SELECT * FROM {tn} WHERE {role} = ?'.format(tn=ref_partner_list, role=sql_lead)
        couple = cursor.execute(query, (identifier,)).fetchall()
        role = ''
        if len(couple) == 0:
            query = 'SELECT * FROM {tn} WHERE {role} = ?'.format(tn=ref_partner_list, role=sql_follow)
            couple = cursor.execute(query, (identifier,)).fetchall()
            if len(couple) != 0:
                role = roles['follow']
        elif len(couple) != 0:
            role = roles['lead']
        couple = couple[0]
        couple_id = couple[partner_dict[sql_no]]
        if role == roles['lead']:
            query = 'UPDATE {tn} SET {role} = "", {city} = "", {ballroom_level} = "", ' \
                    '{latin_level} = "" WHERE {role} = ?' \
                .format(tn=ref_partner_list,
                        role=sql_lead, city=sql_city_lead,
                        ballroom_level=sql_ballroom_level_lead, latin_level=sql_latin_level_lead)
            cursor.execute(query, (identifier,))
        elif role == roles['follow']:
            query = 'UPDATE {tn} SET {role} = "", {city} = "", {ballroom_level} = "", ' \
                    '{latin_level} = "" WHERE {role} = ?' \
                .format(tn=ref_partner_list,
                        role=sql_follow, city=sql_city_follow,
                        ballroom_level=sql_ballroom_level_follow, latin_level=sql_latin_level_follow)
            cursor.execute(query, (identifier,))
        connection.commit()
        query = 'SELECT * FROM {tn} WHERE {num} = ?'.format(tn=ref_partner_list, num=sql_no)
        couple = cursor.execute(query, (couple_id,)).fetchall()
        couple = couple[0]
        if couple[partner_dict[sql_lead]] == '' and couple[partner_dict[sql_follow]] == '':
            query = 'DELETE FROM {tn} WHERE {num} = ?'.format(tn=ref_partner_list, num=sql_no)
            cursor.execute(query, (couple_id,))
            connection.commit()
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
            .format(tn1=selection_list, tn2=signup_list, id=sql_id)
        cursor.execute(query, (identifier,))
        query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=selected_list, id=sql_id)
        cursor.execute(query, (identifier,))
        connection.commit()
        status_print('Removed {} from the NTDS selection.'.format(identifier))


def delete_selected_contestant(identifier, connection, cursor):
    """"Temp"""
    if identifier is not None:
        status_print('Contestant {num} cancelled his/her signup for the NTDS.'.format(num=identifier))
        remove_selected_contestant(identifier, connection=connection, cursor=cursor)
        query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=selected_list, id=sql_id)
        cursor.execute(query, (identifier,))
        query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=selection_list, id=sql_id)
        cursor.execute(query, (identifier,))
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;' \
            .format(tn1=cancelled_list, tn2=signup_list, id=sql_id)
        cursor.execute(query, (identifier,))
        connection.commit()


def reinstate_selected_contestant(identifier, connection, cursor):
    """"Temp"""
    if identifier is not None:
        status_print('Put 0contestant {num} back on the list to be eligible for selection.'.format(num=identifier))
        query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=cancelled_list, id=sql_id)
        cursor.execute(query, (identifier,))
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;' \
            .format(tn1=selection_list, tn2=signup_list, id=sql_id)
        cursor.execute(query, (identifier,))
        connection.commit()


def select_bulk(limit, connection, cursor, no_partner=False):
    """"Temp"""
    if limit > max_contestants:
        limit = max_fixed_lion_contestants
    query = 'SELECT * FROM {tn}'.format(tn=selection_list)
    available_dancers = cursor.execute(query).fetchall()
    number_of_available_dancers = len(available_dancers)
    if number_of_available_dancers > 0:
        random_order = random.sample(range(0, number_of_available_dancers), number_of_available_dancers)
        for num in range(len(random_order)):
            dancer = available_dancers[random_order[num]]
            dancer_id = dancer[gen_dict[sql_id]]
            if dancer_id is not None:
                query = ' SELECT * FROM {tn} WHERE {id} = ?'.format(tn=selection_list, id=sql_id)
                dancer_available = cursor.execute(query, (dancer_id,)).fetchone()
                if dancer_available is not None:
                    partner_id = find_partner(dancer_id, connection=connection, cursor=cursor)
                    if (partner_id is not None and no_partner is False) or no_partner is True:
                        create_pair(dancer_id, partner_id, connection=connection, cursor=cursor)
                        move_selected_contestant(dancer_id, connection=connection, cursor=cursor)
                        move_selected_contestant(partner_id, connection=connection, cursor=cursor)
            query = 'SELECT * FROM {tn}'.format(tn=selected_list)
            number_of_selected_dancers = len(cursor.execute(query).fetchall())
            if number_of_selected_dancers >= limit:
                break
    connection.commit()
    reset_selection_tables(connection=connection, cursor=cursor)


def create_city_beginners_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn1} WHERE ({ballroom_level} = ? OR {latin_level} = ?) AND {team} = ?' \
            .format(tn1=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level, team=sql_city)
        max_city_beginners = len(cursor.execute(query, (levels['beg'], levels['beg'], city)).fetchall())
        if max_city_beginners > max_guaranteed_beginners:
            max_city_beginners = max_guaranteed_beginners
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=fixed_beginners_list)
        cursor.execute(query, (city, 0, max_city_beginners))
    connection.commit()


def update_city_beginners(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ?' \
            .format(tn=partners_list, city_lead=sql_city_lead)
        number_of_city_beginners = len(cursor.execute(query, (city,)).fetchall())
        query = 'SELECT * FROM {tn} WHERE {city_follow} LIKE ?' \
            .format(tn=partners_list, city_follow=sql_city_follow)
        number_of_city_beginners += len(cursor.execute(query, (city,)).fetchall())
        query = 'UPDATE {tn} SET {num} = ? WHERE {city} = ?' \
            .format(tn=fixed_beginners_list, num=sql_num_con, city=sql_city)
        cursor.execute(query, (number_of_city_beginners, city))
    connection.commit()


def get_lions_query():
    """"Temp"""
    query = 'SELECT * FROM {tn1} WHERE {team} = ?' \
        .format(tn1=selection_list, team=sql_city)
    sql_filter = []
    for level in lion_participants:
        sql_filter.append(' ( {ballroom_level} = "' + level + '"' + ' OR {latin_level} = "' + level + '" )')
    query_extension = ' AND (' + ' OR '.join(sql_filter) + ' )'
    query_extension = query_extension.format(ballroom_level=sql_ballroom_level, latin_level=sql_latin_level)
    query += query_extension
    return query


def create_city_lions_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = get_lions_query()
        max_city_lions = len(cursor.execute(query, (city,)).fetchall())
        if max_city_lions > max_fixed_lion_contestants:
            max_city_lions = max_fixed_lion_contestants
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=fixed_lions_list)
        cursor.execute(query, (city, 0, max_city_lions))
    connection.commit()


def update_city_lions(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn} WHERE {city_lead} = ? '.format(tn=partners_list, city_lead=sql_city_lead)
        sql_filter = []
        for lvl in lion_participants:
            sql_filter.append('{ballroom_level_lead} = "{lvl}"'
                              .format(ballroom_level_lead=sql_ballroom_level_lead, lvl=lvl))
        sql_filter.append('{ballroom_level_lead} LIKE "%"'.format(ballroom_level_lead=sql_ballroom_level_lead))
        query += ' AND (' + ' OR '.join(sql_filter) + ')'
        sql_filter = []
        for lvl in lion_participants:
            sql_filter.append('{ballroom_level_follow} = "{lvl}"'
                              .format(ballroom_level_follow=sql_ballroom_level_follow, lvl=lvl))
        sql_filter.append('{ballroom_level_follow} LIKE "%"'.format(ballroom_level_follow=sql_ballroom_level_follow))
        query += ' AND (' + ' OR '.join(sql_filter) + ' )'
        number_of_city_lions = len(cursor.execute(query, (city,)).fetchall())
        #
        query = 'SELECT * FROM {tn} WHERE {city_follow} = ? '.format(tn=partners_list, city_follow=sql_city_follow)
        sql_filter = []
        for lvl in lion_participants:
            sql_filter.append('{ballroom_level_lead} = "{lvl}"'
                              .format(ballroom_level_lead=sql_ballroom_level_lead, lvl=lvl))
        sql_filter.append('{ballroom_level_lead} LIKE "%"'.format(ballroom_level_lead=sql_ballroom_level_lead))
        query += ' AND (' + ' OR '.join(sql_filter) + ')'
        sql_filter = []
        for lvl in lion_participants:
            sql_filter.append('{ballroom_level_follow} = "{lvl}"'
                              .format(ballroom_level_follow=sql_ballroom_level_follow, lvl=lvl))
        sql_filter.append('{ballroom_level_follow} LIKE "%"'.format(ballroom_level_follow=sql_ballroom_level_follow))
        query += ' AND (' + ' OR '.join(sql_filter) + ' )'
        number_of_city_lions += len(cursor.execute(query, (city,)).fetchall())
        query = 'UPDATE {tn} SET {num} = ? WHERE {city} = ?' \
            .format(tn=fixed_lions_list, num=sql_num_con, city=sql_city)
        cursor.execute(query, (number_of_city_lions, city))
    connection.commit()


def max_rc(direction, worksheet):
    """Finds the maximum number of rows or columns of a worksheet"""
    max_dir = 0
    if direction == 'row':
        while True:
            max_dir += 1
            if worksheet.cell(row=max_dir, column=1).value is None:
                max_dir -= 1
                break
    elif direction == 'col':
        while True:
            max_dir += 1
            if worksheet.cell(row=1, column=max_dir).value is None:
                max_dir -= 1
                break
    return max_dir


def create_tables(connection, cursor):
    """"Drops existing tables and creates new ones"""
    # Drop all existing tables (from previous run)
    tables_to_drop = [signup_list, selection_list, selected_list, cancelled_list, backup_list,
                      team_list, partners_list, ref_partner_list, fixed_beginners_list, fixed_lions_list]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        cursor.execute(query)
    # Create new tables
    dancer_list_tables = [signup_list, selection_list, selected_list, cancelled_list, backup_list]
    for item in dancer_list_tables:
        query = dancers_list_query.format(item)
        cursor.execute(query)
    query = team_list_query.format(team_list)
    cursor.execute(query)
    query = paren_table_query.format(partners_list)
    cursor.execute(query)
    query = paren_table_query.format(ref_partner_list)
    cursor.execute(query)
    query = city_list_query.format(fixed_beginners_list)
    cursor.execute(query)
    query = city_list_query.format(fixed_lions_list)
    cursor.execute(query)
    connection.commit()


def create_competing_teams(connection, cursor):
    """"Creates list of all competing cities"""
    # Empty list that will contain the signup sheet file names for each of the teams
    competing_cities_array = []
    # Reads file containing the participating teams
    team_parser = configparser.ConfigParser()
    team_parser.read(participating_teams_key['path'])
    sections = team_parser.sections()
    # Write signup sheets into database
    for section in sections:
        team = team_parser[section]
        competing_cities_array.append([team['TeamName'], team['City'], team['SignupSheet']])
    query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=team_list)
    for row in competing_cities_array:
        if os.path.isfile(path=row[team_dict[sql_signup_list]]):
            cursor.execute(query, row)
    connection.commit()
    query = 'SELECT * FROM {tn} ORDER BY {team}'.format(tn=team_list, team=sql_team)
    competing_cities_array = cursor.execute(query).fetchall()
    return competing_cities_array


def get_competing_teams(cursor):
    """"Temp"""
    query = 'SELECT * FROM {tn} ORDER BY {team}'.format(tn=team_list, team=sql_team)
    return cursor.execute(query).fetchall()


def get_competing_cities(cursor):
    """"Temp"""
    competing_teams = get_competing_teams(cursor)
    competing_cities = [row[team_dict[sql_city]] for row in competing_teams]
    return competing_cities


def reset_selection_tables(connection, cursor):
    """"Resets the partners_list table"""
    query = drop_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()
    query = paren_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()


def collect_city_overview(source_table, target_table, users, cursor, connection, collect_data=False):
    """"Prints an overview"""
    if source_table == selected_list:
        query = 'SELECT {city}, COUNT() FROM {tn} GROUP BY {city}'.format(tn=source_table, city=sql_city)
    else:
        query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=source_table, city=sql_city)
    ordered_cities = cursor.execute(query).fetchall()
    if gather_stats and collect_data:
        query = 'CREATE TABLE IF NOT EXISTS {tn} (id INTEGER PRIMARY KEY AUTOINCREMENT'.format(tn=target_table)
        for city in ordered_cities:
            query += ', {city} INT'.format(city=city[0])
        query += ')'
        cursor.execute(query)
        query = 'INSERT INTO {tn} ('.format(tn=target_table)
        for city in ordered_cities:
            query += '{city}, '.format(city=city[0])
        query = query[:-2] + ') VALUES (' + '?,'*len(ordered_cities)
        query = query[:-1] + ')'
        cursor.execute(query, tuple([city[1] for city in ordered_cities]))
    for city in ordered_cities:
        overview = 'Number of selected {users} from {city} is: {number}'\
            .format(city=city[0], number=city[1], users=users)
        status_print(overview)
    status_print('')
    connection.commit()


def export_excel_lists(cursor, timestamp, city=None):
    """"Creates an Excel file, containing a list of who got selected and who is on the waiting list.
    One file containing all the dancers is generated for the for the organization, as well as a separate for
    each cities' team captain, containing lists of only that city"""
    status_print('Exporting all selected dancers')
    if city is None:
        output_file = output_key['path'] + '_' + timestamp + '_Overview' + xlsx_ext
    else:
        output_file = output_key['path'] + '_' + timestamp + '_' + city + xlsx_ext
    workbook = openpyxl.Workbook()
    worksheet = workbook.worksheets[0]
    worksheet.title = 'Selected'
    output_data = list()
    output_data.append(output_title)
    if city is None:
        query = 'SELECT * FROM {tn} ORDER BY {id}'.format(tn=selected_list, id=sql_id)
        selected_dancers = cursor.execute(query).fetchall()
    else:
        query = 'SELECT * FROM {tn} WHERE {city} = ? ORDER BY {id}'.format(tn=selected_list, city=sql_city, id=sql_id)
        selected_dancers = cursor.execute(query, (city,)).fetchall()
    output_data.extend(selected_dancers)
    for row in range(len(output_data)):
        for column in range(len(output_title)):
            cell = worksheet.cell(row=row + 1, column=column + 1)
            cell.value = output_data[row][column]
    workbook.save(output_file)
    status_print('Exporting all dancers on the waiting list')
    worksheet = workbook.create_sheet('Waiting list')
    output_data = list()
    output_data.append(output_title)
    if city is None:
        query = 'SELECT * FROM {tn} ORDER BY {id}'.format(tn=selection_list, id=sql_id)
        waiting_dancers = cursor.execute(query).fetchall()
    else:
        query = 'SELECT * FROM {tn} WHERE {city} = ? ORDER BY {id}'.format(tn=selection_list, city=sql_city, id=sql_id)
        waiting_dancers = cursor.execute(query, (city,)).fetchall()
    output_data.extend(waiting_dancers)
    for row in range(len(output_data)):
        for column in range(len(output_title)):
            cell = worksheet.cell(row=row + 1, column=column + 1)
            cell.value = output_data[row][column]
    status_print('Saving output: "{file}"'.format(file=output_file))
    status_print('')
    workbook.save(output_file)


def create_stats_file(cursor, timestamp):
    """"Temp"""
    status_print('Exporting statistics file')
    query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=team_list, city=sql_city)
    stats_title = cursor.execute(query).fetchall()
    stats_title = [city[team_dict[sql_city]] for city in stats_title]
    stats_title.insert(0, 'id')
    output_file = output_key['path'] + '_' + timestamp + '_Statistics' + xlsx_ext
    workbook = openpyxl.Workbook()
    worksheet = workbook.worksheets[0]
    worksheet.title = levels['beg']
    query = 'SELECT * FROM {tn} ORDER by "ïd"'.format(tn=beginners_list)
    output_data = list()
    output_data.append(stats_title)
    output_data.extend(cursor.execute(query).fetchall())
    status_print('Exporting {}.'.format(levels['beg']))
    for row in range(len(output_data)):
        for column in range(len(output_title)):
            cell = worksheet.cell(row=row + 1, column=column + 1)
            cell.value = output_data[row][column]
    workbook.save(output_file)
    worksheet = workbook.create_sheet(lions)
    query = 'SELECT * FROM {tn} ORDER by "ïd"'.format(tn=lions_list)
    output_data = list()
    output_data.append(stats_title)
    output_data.extend(cursor.execute(query).fetchall())
    status_print('Exporting {}.'.format(lions))
    for row in range(len(output_data)):
        for column in range(len(output_title)):
            cell = worksheet.cell(row=row + 1, column=column + 1)
            cell.value = output_data[row][column]
    workbook.save(output_file)
    worksheet = workbook.create_sheet(contestants)
    query = 'SELECT * FROM {tn} ORDER by "ïd"'.format(tn=contestants_list)
    output_data = list()
    output_data.append(stats_title)
    output_data.extend(cursor.execute(query).fetchall())
    status_print('Exporting {}.'.format(contestants))
    for row in range(len(output_data)):
        for column in range(len(output_title)):
            cell = worksheet.cell(row=row + 1, column=column + 1)
            cell.value = output_data[row][column]
    workbook.save(output_file)
    # worksheet = workbook.create_sheet(individuals)
    # query = 'SELECT * FROM {tn} ORDER by "ïd"'.format(tn=individual_list)
    # output_data = list()
    # output_data.append(stats_title)
    # output_data.extend(cursor.execute(query).fetchall())
    # status_print('Exporting {}.'.format(individuals))
    # for row in range(len(output_data)):
    #     for column in range(len(output_title)):
    #         cell = worksheet.cell(row=row + 1, column=column + 1)
    #         cell.value = output_data[row][column]
    # workbook.save(output_file)
    status_print('Finished exporting statistics')


def print_ntds_config():
    """"Displays the boundary conditions of the tournament in the welcome text"""
    if user_boundaries:
        status_print('Using user setting for boundary conditions.')
    else:
        status_print('Using default settings.')
    status_print('')
    status_print('Contestant numbers for this NTDS selection:')
    status_print('')
    status_print('Maximum number of contestants: {num}'.format(num=max_contestants))
    status_print('Guaranteed beginners per team: {num}'.format(num=max_guaranteed_beginners))
    status_print('Guaranteed Lion contestants per team: {num}'.format(num=max_fixed_lion_contestants))
    status_print('Cutoff for selecting all beginners: {num}'.format(num=beginner_signup_cutoff))
    status_print('Buffer for selecting contestants at the end: {num}'.format(num=buffer_for_selection))
    msg = 'Levels participating for the Lion: '
    for level in lion_participants:
        msg += '{lvl}, '.format(lvl=level)
    msg = msg[:-2]
    status_print(msg)
    status_print('')


def welcome_text():
    """"Text displayed when opening the program for the first time."""
    status_text.config(state=NORMAL)
    status_text.config(wrap=WORD)
    status_text.config(state=DISABLED)
    status_print('Welcome to the NTDS Selection!')
    status_print('')
    status_print('You can start a new selection, update an existing selection, '
                 'or change the settings with the buttons in the bottom right corner.')
    status_print('')
    status_print('For an overview of the available commands, type "help" in the command prompt.')
    status_print('')
    status_text.config(state=NORMAL)
    status_text.config(wrap=NONE)
    status_text.config(state=DISABLED)
    data_text.config(state=NORMAL)
    data_text.delete('1.0', END)
    data_text.insert(END, 'Data about the number of contestants, First Aid Officers, required sleeping locations, etc. '
                          'will be displayed here once a database has been created/selected.')
    data_text.config(state=DISABLED)


def command_help_text():
    """"Help text"""
    status_print('')
    status_print('Listing all commands:')
    status_print('')
    status_print(list_selected)
    status_print('Lists all dancers that were selected for the tournament.')
    status_print('')
    status_print(list_selected + '_beginners / _breiten / _closed / _open')
    status_print('Lists all dancers of the chosen level that were selected for the tournament.')
    status_print('')
    status_print(list_available)
    status_print('Lists all dancers available for selection.')
    status_print('')
    status_print(list_available + '_beginners / _breiten / _closed / _open')
    status_print('Lists all dancers of the chosen level that are available for selection for the tournament.')
    status_print('')
    status_print(list_backup)
    status_print('Lists all dancers that are on the backup list for the tournament')
    status_print('')
    status_print(list_backup + '_beginners / _breiten / _closed / _open')
    status_print('Lists all dancers of the chosen level that are on the backup list for the tournament.')
    status_print('')
    status_print(list_cancelled)
    status_print('Lists all dancers that have cancelled their signup for the tournament.')
    status_print('')
    status_print(list_cv)
    status_print('Lists all dancers available for selection that want to volunteer at the tournament.')
    status_print('')
    status_print(list_pv)
    status_print('Lists all dancers available for selection that were a volunteer at a previous tournament.')
    status_print('')
    status_print(list_fa)
    status_print('Lists all dancers available for selection that are a qualified First Aid Officer.')
    status_print('')
    status_print(list_ero)
    status_print('Lists all dancers available for selection that are a qualified Emergency Response Officer.')
    status_print('')
    status_print(list_ballroom_jury)
    status_print('Lists all dancers available for selection that can be a Ballroom Jury.')
    status_print('')
    status_print(list_latin_jury)
    status_print('Lists all dancers available for selection that can be a Latin Jury.')
    status_print('')
    status_print(select + ' n')
    status_print('Selects contestant number "n" (and their signed partner) for the NTDS.')
    status_print('')
    status_print(selectp + ' n')
    status_print('Selects contestant number "n", and a (virtual) partner for the NTDS.')
    status_print('')
    status_print(remove + ' n')
    status_print('Removes contestant number "n" from the selected contestants.')
    status_print('')
    status_print(removep + ' n')
    status_print('Removes contestant number "n", and their (virtual) partner from the selected contestants.')
    status_print('')
    status_print(delete + ' n')
    status_print('Cancels the signup of contestant number "n", removing them from the tournament selection pool.')
    status_print('')
    status_print(reinstate + ' n')
    status_print('Reinstates the signup of contestant number "n", placing them into the tournament selection pool.')
    status_print('')
    status_print(selectr + ' input')
    status_print('Selects a random contestant and a (virtual) partner for the NTDS.')
    status_print('"input" can be one of the following statements:')
    status_print('A city name. Doing this will select a random contestant from the given city.')
    status_print(' / '.join((rinput_fa, rinput_ero, rinput_ballroom_jury, rinput_latin_jury)))
    status_print('Select a random dancers that can volunteer as a First Aid Officer, an Emergency Response Officer, '
                 'a Ballroom jury or a Latin jury respectively.')
    status_print('')
    status_print(print_contestants)
    status_print('Prints a list of how much contestants each city has had selected.')
    status_print('')
    status_print(print_breakdown)
    status_print('Prints a list of how much contestants each city has had selected, '
                 'breaking it down per class and discipline.')
    status_print('')
    status_print(finish_selection)
    status_print('Finishes the NTDS selection by adding random contestants.')
    status_print('')
    status_print(export)
    status_print('Creates Excel files containing the selection data.')
    status_print('')
    status_print(stats + ' n')
    status_print('Simulates the automatic selection process "n" times (default {runs}), '
                 'to gather statistical data about the selection process.'.format(runs=runs))
    status_print('Used to verify that the selection process is fair (random).')
    status_print('WARNING: Doing this might take up to 4 hours, longer if the machine running the simulations is being '
                 'actively used. ')
    status_print('')


def status_print(message, wrap=True):
    """"Prints the message passed to the program screen"""
    status_text.config(state=NORMAL)
    if wrap is True:
        message = textwrap.fill(message, status_text_width)
    status_text.insert(END, message)
    status_text.insert(END, '\n')
    status_text.update()
    status_text.see(END)
    status_text.config(state=DISABLED)


def print_table(table):
    """"Formatting for printing tables"""
    formatted_table = []
    for dancer in table:
        if dancer[gen_dict[sql_ln_prefix]] == '':
            name = dancer[gen_dict[sql_first_name]] + ' ' + dancer[gen_dict[sql_last_name]]
        else:
            name = dancer[gen_dict[sql_first_name]] + ' ' + dancer[gen_dict[sql_ln_prefix]] + ' ' + \
                   dancer[gen_dict[sql_last_name]]
        formatted_dancer = [str(dancer[gen_dict[sql_id]]), name,
                            dancer[gen_dict[sql_ballroom_level]], dancer[gen_dict[sql_latin_level]],
                            str(dancer[gen_dict[sql_ballroom_partner]]), str(dancer[gen_dict[sql_latin_partner]]),
                            dancer[gen_dict[sql_ballroom_role]], dancer[gen_dict[sql_latin_role]],
                            dancer[gen_dict[sql_ballroom_mandatory_blind_date]],
                            dancer[gen_dict[sql_latin_mandatory_blind_date]],
                            dancer[gen_dict[sql_first_aid]], dancer[gen_dict[sql_emergency_response_officer]],
                            dancer[gen_dict[sql_ballroom_jury]], dancer[gen_dict[sql_latin_jury]],
                            dancer[gen_dict[sql_student]], dancer[gen_dict[sql_sleeping_location]],
                            dancer[gen_dict[sql_current_volunteer]], dancer[gen_dict[sql_previous_volunteer]],
                            dancer[gen_dict[sql_city]]]
        formatted_table.append(formatted_dancer)
    status_text.config(wrap=NONE)
    print_table_header = ['Id', 'Name', 'Ballroom level', 'Latin level', 'Ballroom partner', 'Latin partner',
                          'Ballroom role', 'Latin role',
                          'Ballroom mandatory blind date', 'Latin mandatory blind date',
                          'First Aid', 'Emergency Response Officer', 'Ballroom jury', 'Latin jury',
                          'Student', 'Sleeping location', 'Volunteer', 'Past volunteer', 'Team']
    status_print(tabulate(formatted_table, headers=print_table_header, tablefmt='grid'), wrap=False)


def status_update():
    """"Data on the contestants of the active database"""
    status_connection = sql.connect(database_key['path'])
    status_cursor = status_connection.cursor()
    query = 'SELECT * FROM {tn}'.format(tn=selected_list)
    number_of_contestants = len(status_cursor.execute(query).fetchall())
    query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {ballroom_role} = ?'\
        .format(tn=selected_list, ballroom_level=sql_ballroom_level, ballroom_role=sql_ballroom_role)
    number_of_beginner_ballroom_leads = len(status_cursor.execute(query, (levels['beg'], roles['lead'])).fetchall())
    number_of_breiten_ballroom_leads = len(status_cursor.execute(query, (levels['breiten'], roles['lead'])).fetchall())
    number_of_closed_ballroom_leads = len(status_cursor.execute(query, (levels['closed'], roles['lead'])).fetchall())
    number_of_open_ballroom_leads = len(status_cursor.execute(query, (levels['open'], roles['lead'])).fetchall())
    number_of_beginner_ballroom_follows = len(status_cursor.execute(query, (levels['beg'], roles['follow'])).fetchall())
    number_of_breiten_ballroom_follows = len(status_cursor.execute(query, (levels['breiten'], roles['follow']))
                                             .fetchall())
    number_of_closed_ballroom_follows = len(status_cursor.execute(query, (levels['closed'], roles['follow']))
                                            .fetchall())
    number_of_open_ballroom_follows = len(status_cursor.execute(query, (levels['open'], roles['follow'])).fetchall())
    query = 'SELECT * FROM {tn} WHERE {latin_level} = ? AND {latin_role} = ?' \
        .format(tn=selected_list, latin_level=sql_latin_level, latin_role=sql_latin_role)
    number_of_beginner_latin_leads = len(status_cursor.execute(query, (levels['beg'], roles['lead'])).fetchall())
    number_of_breiten_latin_leads = len(status_cursor.execute(query, (levels['breiten'], roles['lead'])).fetchall())
    number_of_closed_latin_leads = len(status_cursor.execute(query, (levels['closed'], roles['lead'])).fetchall())
    number_of_open_latin_leads = len(status_cursor.execute(query, (levels['open'], roles['lead'])).fetchall())
    number_of_beginner_latin_follows = len(status_cursor.execute(query, (levels['beg'], roles['follow'])).fetchall())
    number_of_breiten_latin_follows = len(status_cursor.execute(query, (levels['breiten'], roles['follow'])).fetchall())
    number_of_closed_latin_follows = len(status_cursor.execute(query, (levels['closed'], roles['follow'])).fetchall())
    number_of_open_latin_follows = len(status_cursor.execute(query, (levels['open'], roles['follow'])).fetchall())
    query = 'SELECT * FROM {tn} WHERE {first_aid} = ?'.format(tn=selected_list, first_aid=sql_first_aid)
    number_of_first_aid_yes = len(status_cursor.execute(query, (NTDS_options['first_aid']['yes'],)).fetchall())
    number_of_first_aid_maybe = len(status_cursor.execute(query, (NTDS_options['first_aid']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {first_aid} = ?'\
        .format(tn=selected_list, first_aid=sql_emergency_response_officer)
    number_of_emergency_response_officer_yes = \
        len(status_cursor.execute(query, (NTDS_options['emergency_response_officer']['yes'],)).fetchall())
    number_of_emergency_response_officer_maybe = \
        len(status_cursor.execute(query, (NTDS_options['emergency_response_officer']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE ballroom_level = ? AND ballroom_mandatory_blind_date = ?' \
        .format(tn=selected_list, ballroom_level=sql_ballroom_level,
                ballroom_mandatory_blind_date=sql_ballroom_mandatory_blind_date)
    number_of_mandatory_breiten_ballroom_blind_daters = len(status_cursor.execute(query, (levels['breiten'], yes))
                                                            .fetchall())
    query = 'SELECT * FROM {tn} WHERE latin_level = ? AND latin_mandatory_blind_date = ?' \
        .format(tn=selected_list, latin_level=sql_latin_level,
                latin_mandatory_blind_date=sql_latin_mandatory_blind_date)
    number_of_mandatory_breiten_latin_blind_daters = len(status_cursor.execute(query, (levels['breiten'], yes))
                                                         .fetchall())
    query = 'SELECT * FROM {tn} WHERE {ballroom_jury} = ?'.format(tn=selected_list, ballroom_jury=sql_ballroom_jury)
    number_of_ballroom_jury_yes = len(status_cursor.execute(query, (NTDS_options['ballroom_jury']['yes'],)).fetchall())
    number_of_ballroom_jury_maybe = len(status_cursor.execute(query, (NTDS_options['ballroom_jury']['maybe'],))
                                        .fetchall())
    query = 'SELECT * FROM {tn} WHERE {latin_jury} = ?'.format(tn=selected_list, latin_jury=sql_latin_jury)
    number_of_latin_jury_yes = len(status_cursor.execute(query, (NTDS_options['latin_jury']['yes'],)).fetchall())
    number_of_latin_jury_maybe = len(status_cursor.execute(query, (NTDS_options['latin_jury']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {current_volunteer} = ?'\
        .format(tn=selected_list, current_volunteer=sql_current_volunteer)
    number_of_current_volunteer_yes = \
        len(status_cursor.execute(query, (NTDS_options['current_volunteer']['yes'],)).fetchall())
    number_of_current_volunteer_maybe = \
        len(status_cursor.execute(query, (NTDS_options['current_volunteer']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {past_volunteer} = ?'\
        .format(tn=selected_list, past_volunteer=sql_previous_volunteer)
    number_of_past_volunteer = len(status_cursor.execute(query, (NTDS_options['past_volunteer']['yes'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {sleeping_location} = ?' \
        .format(tn=selected_list, sleeping_location=sql_sleeping_location)
    number_of_sleeping_spots = len(status_cursor.execute(query, (NTDS_options['sleeping_location']['yes'],)).fetchall())
    data_text.config(state=NORMAL)
    data_text.delete('1.0', END)
    data_text.insert(END, 'Selected database name: {name}\n'.format(name=database_key['path']))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Total number of contestants: {num}\n'.format(num=number_of_contestants))
    data_text.insert(END, '\n')
    if levels['beg'] in classes:
        data_text.insert(END, 'Beginners Ballroom Leads: {num}\n'.format(num=number_of_beginner_ballroom_leads))
        data_text.insert(END, 'Beginners Ballroom Follows: {num}\n'.format(num=number_of_beginner_ballroom_follows))
        data_text.insert(END, 'Beginners Latin Leads: {num}\n'.format(num=number_of_beginner_latin_leads))
        data_text.insert(END, 'Beginners Latin Follows: {num}\n'.format(num=number_of_beginner_latin_follows))
        data_text.insert(END, '\n')
    if levels['breiten'] in classes:
        data_text.insert(END, 'Breitensport Ballroom Leads: {num}\n'.format(num=number_of_breiten_ballroom_leads))
        data_text.insert(END, 'Breitensport Ballroom Follows: {num}\n'.format(num=number_of_breiten_ballroom_follows))
        data_text.insert(END, 'Breitensport Latin Leads: {num}\n'.format(num=number_of_breiten_latin_leads))
        data_text.insert(END, 'Breitensport Latin Follows: {num}\n'.format(num=number_of_breiten_latin_follows))
        data_text.insert(END, '\n')
    if levels['closed'] in classes:
        data_text.insert(END, 'CloseD Ballroom Leads: {num}\n'.format(num=number_of_closed_ballroom_leads))
        data_text.insert(END, 'CloseD Ballroom Follows: {num}\n'.format(num=number_of_closed_ballroom_follows))
        data_text.insert(END, 'CloseD Latin Leads: {num}\n'.format(num=number_of_closed_latin_leads))
        data_text.insert(END, 'CloseD Latin Follows: {num}\n'.format(num=number_of_closed_latin_follows))
        data_text.insert(END, '\n')
    if levels['open'] in classes:
        data_text.insert(END, 'Open Class Ballroom Leads: {num}\n'.format(num=number_of_open_ballroom_leads))
        data_text.insert(END, 'Open Class Ballroom Follows: {num}\n'.format(num=number_of_open_ballroom_follows))
        data_text.insert(END, 'Open Class Latin Leads: {num}\n'.format(num=number_of_open_latin_leads))
        data_text.insert(END, 'Open Class Latin Follows: {num}\n'.format(num=number_of_open_latin_follows))
        data_text.insert(END, '\n')
    data_text.insert(END, 'First Aid: {yes} ({maybe})\n'
                     .format(yes=number_of_first_aid_yes, maybe=number_of_first_aid_maybe))
    data_text.insert(END, 'Emergency Response Officer: {yes} ({maybe})\n'
                     .format(yes=number_of_emergency_response_officer_yes,
                             maybe=number_of_emergency_response_officer_maybe))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Breitensport mandatory Ballroom blind daters: {num}\n'
                     .format(num=number_of_mandatory_breiten_ballroom_blind_daters))
    data_text.insert(END, 'Breitensport mandatory Latin blind daters: {num}\n'
                     .format(num=number_of_mandatory_breiten_latin_blind_daters))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Ballroom juries: {yes} ({maybe})\n'
                     .format(yes=number_of_ballroom_jury_yes, maybe=number_of_ballroom_jury_maybe))
    data_text.insert(END, 'Latin juries {yes} ({maybe})\n'
                     .format(yes=number_of_latin_jury_yes, maybe=number_of_latin_jury_maybe))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Volunteers: {yes} ({maybe})\n'
                     .format(yes=number_of_current_volunteer_yes, maybe=number_of_current_volunteer_maybe))
    data_text.insert(END, 'Past volunteers: {yes}\n'.format(yes=number_of_past_volunteer))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Sleeping spots: {yes}'.format(yes=number_of_sleeping_spots))
    data_text.see(END)
    data_text.config(state=DISABLED)
    status_cursor.close()
    status_connection.close()


def select_database(entry=None):
    """"Temp"""
    old_database_name = database_key['db']
    if entry is None:
        ask_database = EntryBox('Please give the database name', (database_key, 'db'))
        root.wait_window(ask_database.top)
    else:
        database_key['db'] = entry
    database_key['path'] = database_key['db'] + db_ext
    if os.path.isfile(path=database_key['path']) and old_database_name != database_key['db']:
        status_text.config(state=NORMAL)
        status_text.delete('1.0', END)
        status_text.config(state=DISABLED)
        status_print('Selected existing database: {name}'.format(name=database_key['path']))
        status_print('')
        cli_text.focus_set()
        sel_conn = sql.connect(database_key['path'])
        sel_curs = sel_conn.cursor()
        status_update()
        sel_curs.close()
        sel_conn.close()
    elif os.path.isfile(path=database_key['path']) is False or old_database_name != database_key['db']:
        status_print('"{name}" is not a valid database.'.format(name=database_key['path']))


def options_menu():
    """"Temp"""
    status_print('Work in progress...')


def list_list(table_list, cursor):
    query = 'SELECT * FROM {tn} ORDER BY {id}'.format(tn=table_list, id=sql_id)
    selected_contestants = cursor.execute(query).fetchall()
    if len(selected_contestants) > 0:
        status_print('')
        if table_list == selection_list:
            status_print('All {num} contestants that are available for selection for the NTDS:'
                         .format(num=len(selected_contestants)))
        elif table_list == selected_list:
            status_print('All {num} contestants that have been selected for the NTDS:'
                         .format(num=len(selected_contestants)))
        elif table_list == cancelled_list:
            status_print('All {num} contestants that have cancelled their signup for the NTDS:'
                         .format(num=len(selected_contestants)))
        elif table_list == backup_list:
            status_print('All {num} contestants that that are on the backup list for the NTDS:'
                         .format(num=len(selected_contestants)))
        status_print('')
        print_table(selected_contestants)
        status_print('')
    else:
        if table_list == selection_list:
            status_print('There are no contestants available for selection for the NTDS.')
        elif table_list == selected_list:
            status_print('There are no contestants selected for the NTDS.')
        elif table_list == cancelled_list:
            status_print('There are no contestants that have cancelled their signup for the NTDS.')
        elif table_list == backup_list:
            status_print('There are no contestants on the backup list for the NTDS.')
        status_print('')


def list_level(level, cursor, table_list=selection_list):
    """"Temp"""
    if level in classes:
        query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? OR {latin_level} = ? ORDER BY {id}' \
            .format(tn=table_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level, id=sql_id)
        available_contestants = cursor.execute(query, (level, level)).fetchall()
        if len(available_contestants) > 0:
            status_print('')
            if table_list == selection_list:
                status_print('All {num} {lvl} dancers that are available for selection for the NTDS:'
                             .format(num=len(available_contestants), lvl=level))
            elif table_list == selected_list:
                status_print('All {num} {lvl} dancers that are selected for the NTDS:'
                             .format(num=len(available_contestants), lvl=level))
            elif table_list == backup_list:
                status_print('All {num} {lvl} dancers that are on the backup list for the NTDS:'
                             .format(num=len(available_contestants), lvl=level))
            status_print('')
            print_table(available_contestants)
            status_print('')
        else:
            if table_list == selection_list:
                status_print('There are no {lvl} dancers available for selection for the NTDS'.format(lvl=level))
            elif table_list == selected_list:
                status_print('There are no {lvl} dancers that have been selected for the NTDS'.format(lvl=level))
            elif table_list == backup_list:
                status_print('There are no {lvl} dancers on the backup list for the NTDS'.format(lvl=level))
            status_print('')
    else:
        status_print('The {lvl} class is not participating in this tournament.'.format(lvl=level))


def list_volunteer(volunteer_role, cursor):
    """"Temp"""
    if volunteer_role == sql_previous_volunteer:
        messages = ['',
                    'Contestants that were a volunteer at a previous tournament:',
                    'There are no contestants available for selection that volunteered at a previous tournament.']
    else:
        messages = ['Contestants that MIGHT want to volunteer ',
                    'Contestants that want to volunteer ',
                    'There are no contestants available for selection that want to volunteer ']
        volunteer_dict = {sql_first_aid: 'as a First Aid Officer:',
                          sql_emergency_response_officer: 'as an Emergency Response Officer:',
                          sql_ballroom_jury: 'as a Ballroom Jury:', sql_latin_jury: 'as a Latin Jury:',
                          sql_current_volunteer: '.'}
        messages = [message + volunteer_dict[volunteer_role] for message in messages]
    query = 'SELECT * FROM {tn} WHERE {v_role} = ? ORDER BY {id}' \
        .format(tn=selection_list, v_role=volunteer_role, id=sql_id)
    available_contestants = cursor.execute(query, (options_ymn['maybe'],)).fetchall()
    if len(available_contestants) > 0:
        status_print('')
        status_print(messages[0])
        status_print('')
        print_table(available_contestants)
        status_print('')
    available_contestants = cursor.execute(query, (options_ymn['yes'],)).fetchall()
    if len(available_contestants) > 0:
        status_print('')
        status_print(messages[1])
        status_print('')
        print_table(available_contestants)
        status_print('')
    else:
        status_print(messages[2])
        status_print('')


def select_random_volunteer(start_mesage, sql_var, cursor):
    """"Temp"""
    status_print(start_mesage)
    query = 'SELECT * from {tn} WHERE {user_input} = ?'.format(tn=selection_list, user_input=sql_var)
    dancers = cursor.execute(query, (yes,)).fetchall()
    if dancers is not None:
        dancer = random.choice(dancers)
        dancer_id = dancer[gen_dict[sql_id]]
        cli_parser('', alternate_input=selectp + ' {id}'.format(id=dancer_id))
    else:
        status_print('There are nog contestants that definitely want to fulfill this role.')
        status_print('Looking for contestants that might want to fulfill it.')
        dancers = cursor.execute(query, (maybe,)).fetchall()
        if dancers is not None:
            dancer = random.choice(dancers)
            dancer_id = dancer[gen_dict[sql_id]]
            cli_parser('', alternate_input=selectp + ' {id}'.format(id=dancer_id))
        else:
            status_print('There are no volunteers for this role to select')


# CLI commands
# Open commands, always available
echo = 'echo'
help_com = 'help'
exit_com = 'exit'
stats = '-stats'
db = '-db'
# Database commands, only available when a database is selected
select = '-select'
selectp = '-selectp'
remove = '-remove'
removep = '-removep'
delete = '-delete'
reinstate = '-reinstate'
selectr = '-selectr'
list_selected = 'list_selected'
list_available = 'list_available'
list_cancelled = 'list_cancelled'
list_backup = 'list_backup'
list_selected_beginners = 'list_selected_beginners'
list_selected_breiten = 'list_selected_breiten'
list_selected_closed = 'list_selected_closed'
list_selected_open = 'list_selected_open'
list_available_beginners = 'list_available_beginners'
list_available_breiten = 'list_available_breiten'
list_available_closed = 'list_available_closed'
list_available_open = 'list_available_open'
list_backup_beginners = 'list_backup_beginners'
list_backup_breiten = 'list_backup_breiten'
list_backup_closed = 'list_backup_closed'
list_backup_open = 'list_backup_open'
list_cv = 'list_cv'
list_pv = 'list_pv'
list_fa = 'list_fa'
list_ero = 'list_ero'
list_ballroom_jury = 'list_ballroom_jury'
list_latin_jury = 'list_latin_jury'
print_contestants = 'print_contestants'
print_breakdown = 'print_breakdown'
finish_selection = 'finish_selection'
export = 'export'

rinput_fa = 'first_aid'
rinput_ero = 'ERO'
rinput_ballroom_jury = 'ballroom_jury'
rinput_latin_jury = 'latin_jury'

def cli_parser(event, alternate_input=''):
    """"Command line interface parser. Reads the given input and acts accordingly."""
    user_input = ''
    if alternate_input == '':
        command = cli_text.get()
        cli_text.delete(0, END)
    else:
        command = alternate_input

    connection = sql.connect(database_key['path'])
    cli_curs = connection.cursor()
    open_commands = [echo, help_com, exit_com,
                     stats, db]
    db_commands = [select, selectp, remove, removep, delete, reinstate, selectr,
                   list_selected, list_available, list_cancelled, list_backup,
                   list_selected_beginners, list_selected_breiten, list_selected_closed, list_selected_open,
                   list_available_beginners, list_available_breiten, list_available_closed, list_available_open,
                   list_backup_beginners, list_backup_breiten, list_backup_closed, list_backup_open,
                   list_cv, list_pv, list_fa, list_ero, list_ballroom_jury, list_latin_jury,
                   print_contestants, print_breakdown,
                   finish_selection, export]
    if command.startswith('-'):
        command = command.split(' ', 1)
        if len(command) > 1:
            user_input = command[1]
        else:
            user_input = ''
        command = command[0]

    if command in db_commands and os.path.isfile(database_key['path']) is True:
        if command == select:
            try:
                selected_id = int(user_input)
                query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
                dancer = cli_curs.execute(query, (selected_id,)).fetchone()
                if dancer is not None:
                    ballroom_partner = dancer[gen_dict[sql_ballroom_partner]]
                    latin_partner = dancer[gen_dict[sql_latin_partner]]
                    if ballroom_partner != '' or latin_partner != '':
                        status_print('Dancer {num} signed with a partner, selecting both'.format(num=selected_id))
                        partner_id = find_partner(selected_id, connection=connection, cursor=cli_curs,
                                                  signed_partner_only=True)
                        create_pair(selected_id, partner_id, connection=connection, cursor=cli_curs)
                        move_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                        move_selected_contestant(partner_id, connection=connection, cursor=cli_curs)
                    else:
                        status_print('Selecting dancer number {num} on his/her own'.format(num=selected_id))
                        create_pair(selected_id, '', connection=connection, cursor=cli_curs)
                        move_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                else:
                    status_print('Dancer {num} is not on the selection list, '
                                 'and can therefor not be selected for the tournament'.format(num=selected_id))
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == selectp:
            try:
                selected_id = int(user_input)
                query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
                dancer = cli_curs.execute(query, (selected_id,)).fetchone()
                if dancer is not None:
                    partner_id = find_partner(selected_id, connection=connection, cursor=cli_curs)
                    create_pair(selected_id, partner_id, connection=connection, cursor=cli_curs)
                    move_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                    move_selected_contestant(partner_id, connection=connection, cursor=cli_curs)
                else:
                    status_print('Dancer {num} is not on the selection list, '
                                 'and can therefor not be selected for the tournament'.format(num=selected_id))
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == remove:
            try:
                selected_id = int(user_input)
                query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selected_list, id=sql_id)
                dancer = cli_curs.execute(query, (selected_id,)).fetchone()
                if dancer is not None:
                    remove_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                else:
                    status_print('Dancer {num} is not on the selected list, '
                                 'and can therefor not be removed from the tournament.'.format(num=selected_id))
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == removep:
            try:
                selected_id = int(user_input)
                query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selected_list, id=sql_id)
                dancer = cli_curs.execute(query, (selected_id,)).fetchone()
                if dancer is not None:
                    partner_id = ''
                    query = 'SELECT * FROM {tn} WHERE {role} = ?'.format(tn=ref_partner_list, role=sql_lead)
                    couple = cli_curs.execute(query, (selected_id,)).fetchall()
                    role = ''
                    if len(couple) == 0:
                        query = 'SELECT * FROM {tn} WHERE {role} = ?'.format(tn=ref_partner_list, role=sql_follow)
                        couple = cli_curs.execute(query, (selected_id,)).fetchall()
                        if len(couple) != 0:
                            role = roles['follow']
                    elif len(couple) != 0:
                        role = roles['lead']
                    couple = couple[0]
                    if role == roles['lead']:
                        partner_id = couple[partner_dict[sql_follow]]
                    elif role == roles['follow']:
                        partner_id = couple[partner_dict[sql_lead]]
                    if partner_id != '':
                        status_print('Removing dancers number {sel} and {par} from the NTDS.'
                                     .format(sel=selected_id, par=partner_id))
                    remove_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                    remove_selected_contestant(partner_id, connection=connection, cursor=cli_curs)
                else:
                    status_print('Dancer {num} is not on the selected list, '
                                 'and can therefor not be removed from the tournament.'.format(num=selected_id))
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == delete:
            try:
                selected_id = int(user_input)
                delete_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == reinstate:
            try:
                selected_id = int(user_input)
                query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=cancelled_list, id=sql_id)
                dancer = cli_curs.execute(query, (selected_id,)).fetchone()
                if dancer is not None:
                    reinstate_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
                else:
                    status_print('Dancer {num} is not on the cancelled list, '
                                 'and is already a part of the tournament selection process.'.format(num=selected_id))
            except ValueError:
                status_print('No/Incorrect user number given.')
        elif command == selectr:
            # Create competing cities list
            cities = get_competing_cities(cursor=cli_curs)
            if user_input in cities:
                status_print('Selecting random contestant from ' + user_input + '.')
                query = 'SELECT * from {tn} WHERE {user_input} = ?'.format(tn=selection_list, user_input=sql_city)
                dancers = cli_curs.execute(query, (user_input,)).fetchall()
                if dancers is not None:
                    dancer = random.choice(dancers)
                    dancer_id = dancer[gen_dict[sql_id]]
                    cli_parser('', alternate_input=selectp+' {id}'.format(id=dancer_id))
                else:
                    status_print('There are no dancers from {city} to select'.format(city=user_input))
            elif user_input == rinput_fa:
                select_random_volunteer('Selecting random First Aid volunteer.', sql_first_aid, cursor=cli_curs)
            elif user_input == rinput_ero:
                select_random_volunteer('Selecting random Emergency Response Officer volunteer.',
                                        sql_emergency_response_officer, cursor=cli_curs)
            elif user_input == rinput_ballroom_jury:
                select_random_volunteer('Selecting random Ballroom jury volunteer.', sql_ballroom_jury, cursor=cli_curs)
            elif user_input == rinput_latin_jury:
                select_random_volunteer('Selecting random Latin jury volunteer.', sql_latin_jury, cursor=cli_curs)
            else:
                status_print(user_input + ' is not a valid input.')

        # List contestants (from a certain class) from specific lists
        elif command == list_selected:
            list_list(selected_list, cursor=cli_curs)
        elif command == list_available:
            list_list(selection_list, cursor=cli_curs)
        elif command == list_cancelled:
            list_list(cancelled_list, cursor=cli_curs)
        elif command == list_backup:
            list_list(backup_list, cursor=cli_curs)
        elif command == list_selected_beginners:
            list_level(levels['beg'], cursor=cli_curs, table_list=selected_list)
        elif command == list_selected_breiten:
            list_level(levels['breiten'], cursor=cli_curs, table_list=selected_list)
        elif command == list_selected_closed:
            list_level(levels['closed'], cursor=cli_curs, table_list=selected_list)
        elif command == list_selected_open:
            list_level(levels['open'], cursor=cli_curs, table_list=selected_list)
        elif command == list_available_beginners:
            list_level(levels['beg'], cursor=cli_curs)
        elif command == list_available_breiten:
            list_level(levels['breiten'], cursor=cli_curs)
        elif command == list_available_closed:
            list_level(levels['closed'], cursor=cli_curs)
        elif command == list_available_open:
            list_level(levels['open'], cursor=cli_curs)
        elif command == list_backup_beginners:
            list_level(levels['beg'], cursor=cli_curs, table_list=backup_list)
        elif command == list_backup_breiten:
            list_level(levels['breiten'], cursor=cli_curs, table_list=backup_list)
        elif command == list_backup_closed:
            list_level(levels['closed'], cursor=cli_curs, table_list=backup_list)
        elif command == list_backup_open:
            list_level(levels['open'], cursor=cli_curs, table_list=backup_list)
        # List volunteers (with a specific role)
        elif command == list_cv:
            list_volunteer(sql_current_volunteer, cursor=cli_curs)
        elif command == list_pv:
            list_volunteer(sql_previous_volunteer, cursor=cli_curs)
        elif command == list_fa:
            list_volunteer(sql_first_aid, cursor=cli_curs)
        elif command == list_ero:
            list_volunteer(sql_emergency_response_officer, cursor=cli_curs)
        elif command == list_ballroom_jury:
            list_volunteer(sql_ballroom_jury, cursor=cli_curs)
        elif command == list_latin_jury:
            list_volunteer(sql_latin_jury, cursor=cli_curs)

        # Finish the selection of the tournament automatically
        elif command == finish_selection:
            status_print('')
            status_print('Finishing the selection for the NTDS automatically')
            select_bulk(max_contestants-1, connection=connection, cursor=cli_curs)
            # TODO add smart system to fill in gaps
            query = 'SELECT * FROM {tn}'.format(tn=selected_list)
            number_of_selected_dancers = len(cli_curs.execute(query).fetchall())
            if number_of_selected_dancers < max_contestants:
                select_bulk(max_contestants, connection=connection, cursor=cli_curs, no_partner=True)
        # Print number of contestants for each city
        elif command == print_contestants:
            status_print('')
            collect_city_overview(source_table=selected_list, target_table=contestants_list, users=contestants,
                                  cursor=cli_curs, connection=connection)
        elif command == print_breakdown:
            print_dict = {sql_ballroom_level: 'Ballroom', sql_latin_level: 'Latin'}
            query = 'SELECT {city}, COUNT() FROM {tn} GROUP BY {city}'.format(tn=selected_list, city=sql_city)
            ordered_cities = cli_curs.execute(query, ()).fetchall()
            query_entries = list()
            for item in classes:
                query_entries.append([sql_ballroom_level, item])
                query_entries.append([sql_latin_level, item])
            for city in ordered_cities:
                city = city[0]
                query = 'SELECT * FROM {tn} WHERE {city} = ?'.format(tn=selected_list, city=sql_city)
                number_of_selected_city_dancers = len(cli_curs.execute(query, (city,)).fetchall())
                status_print('')
                status_print('Breakdown of the number of selected dancers from {city}:'.format(city=city))
                status_print('Total dancers selected:\t\t {num}'.format(num=number_of_selected_city_dancers))
                for entries in query_entries:
                    query = 'SELECT * FROM {tn} WHERE {temp} = ? AND {city} = ?' \
                        .format(tn=selected_list, city=sql_city, temp={})
                    query = query.format(entries[0])
                    number_of_selected_city_dancers = len(cli_curs.execute(query, (entries[1], city)).fetchall())
                    # TODO formatting
                    status_print('{level}, {division}:\t\t {num}'
                                 .format(division=print_dict[entries[0]], level=entries[1],
                                         num=number_of_selected_city_dancers))
        elif command == export:
            # Unix timestamp at time of exporting
            save_time = str(round(time.time()))
            # Export database (make copy)
            output_db = output_key['path'] + '_' + save_time + db_ext
            status_print('')
            status_print('Exporting database')
            status_print('Saving output: "{file}"'.format(file=output_db))
            status_print('')
            shutil.copy2(database_key['path'], output_db)
            # Export overview file
            status_print('Exporting overview file')
            export_excel_lists(cursor=cli_curs, timestamp=save_time)
            query = 'SELECT {city} FROM {tn} ORDER BY {city}'.format(tn=team_list, city=sql_city)
            all_cities = cli_curs.execute(query).fetchall()
            for city in all_cities:
                city = city[0]
                status_print('Exporting {city} overview file'.format(city=city))
                export_excel_lists(cursor=cli_curs, timestamp=save_time, city=city)
            status_print('Export complete.')
            status_print('Exported files can be found in the folder:')
            status_print('"{folder}"'.format(folder=output_key['path'].replace(output_key['name'], '')))
    if command in open_commands:
        if command == echo:
            status_print(command)
            create_stats_file(cursor=cli_curs, timestamp=str(round(time.time())))
        elif command == help_com:
            command_help_text()
        elif command == exit_com:
            status_print('')
            status_print('Closing down program.')
            status_print('')
            time.sleep(0.5)
            root.destroy()
        elif command == stats:
            try:
                iterations = int(user_input)
            except ValueError:
                iterations = runs
            global gather_stats
            gather_stats = True
            duration = time.time()
            main_selection()
            if iterations > 1:
                duration = int(time.time() - duration)
                duration_warning = 'One selection run took about {time} seconds.\n'.format(time=duration)
                duration_warning += 'Running all of the selections will take approximately {time} minutes.\n'\
                    .format(time=round(iterations*duration/60))
                duration_warning += 'Proceed?'
                if messagebox.askyesno(root, message=duration_warning):
                    for i in range(iterations-1):
                        status_print('')
                        status_print('Starting run {i} of {it}'.format(i=i+2, it=iterations))
                        status_print('')
                        main_selection()
                    # create_stats_file(cursor=cli_curs, timestamp=str(round(time.time())))
                else:
                    status_print('Cancelling statistics gathering.')
            gather_stats = False
        elif command == db:
            selected_db = str(user_input)
            select_database(entry=selected_db)
    if command != '' and command not in open_commands and command not in db_commands:
        status_print('Unknown command: "{command}"'.format(command=command))
    elif command in db_commands and os.path.isfile(database_key['path']) is False:
        status_print('Command not available, no open database')
    if os.path.isfile(database_key['path']) is True and root.winfo_exists() == 1:
        status_update()
    cli_curs.close()
    connection.close()
    return event


def main_selection():
    ####################################################################################################################
    # Extract data from sign-up sheets
    ####################################################################################################################
    # Disable the CLI
    cli_text.config(state=DISABLED)
    # Time at which the run started
    start_time = time.time()
    # Connect to database and create a cursor
    status_print('')
    status_print('Creating new database...')
    status_print('Database name: {name}'.format(name=default_db_name+db_ext))
    database_key['db'] = default_db_name
    database_key['path'] = database_key['db'] + db_ext
    conn = sql.connect(database_key['path'])
    curs = conn.cursor()
    # Create SQL tables
    status_print('')
    status_print('Creating tables...')
    create_tables(connection=conn, cursor=curs)
    # Create competing cities list
    competing_teams = create_competing_teams(connection=conn, cursor=curs)
    competing_cities = [row[team_dict[sql_city]] for row in competing_teams]
    # Get maximum number of columns
    wb = openpyxl.load_workbook(template_key['path'])
    ws = wb.worksheets[0]
    max_col = max_rc('col', ws)
    # Copy the signup list from every team into the SQL database
    total_signup_list = list()
    total_number_of_contestants = 0
    status_print('')
    status_print('Collecting signup data...')
    for team in competing_teams:
        city = team[team_dict[sql_city]]
        team_signup_list = team[team_dict[sql_signup_list]]
        if os.path.isfile(path=team_signup_list):
            status_print('{city} is entering the tournament, adding contestants to signup list'.format(city=city))
            # Get maximum number of rows and extract signup list
            wb = openpyxl.load_workbook(team_signup_list, data_only=True)
            ws = wb.worksheets[0]
            max_row = max_rc('row', ws)
            city_signup_list = list(ws.iter_rows(min_col=1, min_row=2, max_col=max_col, max_row=max_row))
            # Convert data to 2d list, replace None values with an empty string,
            # increase the id numbers so that there are no duplicates, and add the city to the contestant
            city_signup_list = [[cell.value for cell in row] for row in city_signup_list]
            city_signup_list = [['' if elem is None else elem for elem in row] for row in city_signup_list]
            city_signup_list = [[elem + total_number_of_contestants if isinstance(elem, int) else elem for elem in row]
                                for row in city_signup_list]
            for row in city_signup_list:
                row.append(city)
            total_signup_list.extend(city_signup_list)
            # Copy the contestants to the SQL database
            query = 'INSERT INTO {} VALUES ('.format(signup_list) + ('?,' * (max_col + 1))[:-1] + ');'
            for row in city_signup_list:
                curs.execute(query, row)
            query = 'INSERT INTO {} VALUES ('.format(selection_list) + ('?,' * (max_col + 1))[:-1] + ');'
            for row in city_signup_list:
                curs.execute(query, row)
            total_number_of_contestants += max_row - 1
        else:
            status_print('{city} is not entering the tournament'.format(city=city))
        conn.commit()
    conn.commit()

    ####################################################################################################################
    # Select the team captains and (virtual) partners
    ####################################################################################################################
    status_print('')
    status_print('Selecting team captains...')
    query = 'WITH cities AS (SELECT DISTINCT {city} FROM {tn1} ORDER BY RANDOM()) SELECT {tn1}.* FROM cities ' \
            'LEFT JOIN {tn1} ON ({tn1}.{city} = cities.{city} AND {tn1}.{tc} = "Ja") ORDER BY {tn1}.{tc}'\
        .format(tn1=selection_list, tc=sql_team_captain, city=sql_city)
    team_captains = curs.execute(query).fetchall()
    for captain in team_captains:
        captain_id = captain[gen_dict[sql_id]]
        query = 'SELECT * FROM {tn1} WHERE {id} = ?'.format(tn1=selected_list, id=sql_id)
        captain_selected = curs.execute(query, (captain_id,)).fetchone()
        if captain_selected is None:
            partner_id = find_partner(captain_id, connection=conn, cursor=curs)
            create_pair(captain_id, partner_id, connection=conn, cursor=curs)
            move_selected_contestant(captain_id, connection=conn, cursor=curs)
            move_selected_contestant(partner_id, connection=conn, cursor=curs)
    conn.commit()
    query = drop_table_query.format(partners_list)
    curs.execute(query)
    query = paren_table_query.format(partners_list)
    curs.execute(query)

    ####################################################################################################################
    # Select beginners if less people have signed than the given cutoff
    ####################################################################################################################
    query = 'SELECT * FROM {tn1} WHERE {ballroom_level} = ? OR {latin_level} = ?' \
        .format(tn1=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level)
    all_beginners = curs.execute(query, (levels['beg'], levels['beg'])).fetchall()
    number_of_signed_beginners = len(all_beginners)
    if number_of_signed_beginners <= beginner_signup_cutoff:
        status_print('')
        status_print('Less than {num} Beginners signed up.'.format(num=beginner_signup_cutoff+1))
        status_print('Matching up as much couples as possible and selecting everyone...')
        create_city_beginners_list(competing_cities, connection=conn, cursor=curs)
        for beg in all_beginners:
            beg_id = beg[gen_dict[sql_id]]
            query = 'SELECT * FROM {tn1} WHERE {id} = ?'.format(tn1=selected_list, id=sql_id)
            beginner_selected = curs.execute(query, (beg_id,)).fetchone()
            if beginner_selected is None:
                partner_id = find_partner(beg_id, connection=conn, cursor=curs)
                create_pair(beg_id, partner_id, connection=conn, cursor=curs)
                move_selected_contestant(beg_id, connection=conn, cursor=curs)
                move_selected_contestant(partner_id, connection=conn, cursor=curs)
                update_city_beginners(competing_cities, connection=conn, cursor=curs)
        conn.commit()
        reset_selection_tables(connection=conn, cursor=curs)

    ####################################################################################################################
    # Select beginners if more people have signed than the given cutoff
    ####################################################################################################################
    if number_of_signed_beginners > beginner_signup_cutoff:
        status_print('')
        status_print('More than {num} Beginners signed up.'.format(num=beginner_signup_cutoff))
        status_print('Selecting guaranteed beginners for each team...')
        create_city_beginners_list(competing_cities, connection=conn, cursor=curs)
        query = 'SELECT sum({max_beg})-sum({num}) FROM {tn}'\
            .format(tn=fixed_beginners_list, max_beg=sql_max_con, num=sql_num_con)
        max_iterations = curs.execute(query).fetchone()[0]
        for iteration in range(max_iterations):
            query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_beginners_list, num=sql_num_con)
            ordered_cities = curs.execute(query).fetchall()
            selected_city = None
            for city in ordered_cities:
                if selected_city is None:
                    number_of_selected_city_beginners = city[city_dict[sql_num_con]]
                    max_number_of_selected_city_beginners = city[city_dict[sql_max_con]]
                    if (max_number_of_selected_city_beginners - number_of_selected_city_beginners) > 0:
                        selected_city = city[city_dict[sql_city]]
            if selected_city is not None:
                query = 'SELECT * FROM {tn1} WHERE ({ballroom_level} = ? OR {latin_level} = ?) AND {city} = ?' \
                    .format(tn1=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                            city=sql_city)
                selected_city_beginners = curs.execute(query, (levels['beg'], levels['beg'], selected_city,)).fetchall()
                number_of_city_beginners = len(selected_city_beginners)
                if number_of_city_beginners > 0:
                    random_order = random.sample(range(0, number_of_city_beginners), number_of_city_beginners)
                    partner_id = None
                    for order_city in ordered_cities:
                        if partner_id is None:
                            order_city = order_city[city_dict[sql_city]]
                            if order_city != selected_city:
                                query = 'SELECT * FROM {tn1} WHERE ({ballroom_level} = ? OR {latin_level} = ?) ' \
                                        'AND {city} = ?'\
                                    .format(tn1=selection_list, ballroom_level=sql_ballroom_level,
                                            latin_level=sql_latin_level, city=sql_city)
                                order_city_beginners = curs.execute(query, (levels['beg'], levels['beg'], order_city,))\
                                    .fetchall()
                                number_of_available_beginners = len(order_city_beginners)
                                if number_of_available_beginners > 0:
                                    for num in random_order:
                                        if partner_id is None:
                                            beg = selected_city_beginners[num]
                                            beginner_id = beg[gen_dict[sql_id]]
                                            query = ' SELECT * FROM {tn} WHERE {id} = ?'\
                                                .format(tn=selected_list, id=sql_id)
                                            beginner_available = curs.execute(query, (beginner_id,)).fetchone()
                                            if beginner_available is None:
                                                partner_id = find_partner(beginner_id, connection=conn, cursor=curs,
                                                                          city=order_city)
                                                if partner_id is not None:
                                                    create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                                                    move_selected_contestant(beginner_id, connection=conn, cursor=curs)
                                                    move_selected_contestant(partner_id, connection=conn, cursor=curs)
                                                    update_city_beginners(competing_cities, connection=conn,
                                                                          cursor=curs)
        # Select beginner that were guaranteed entry, but could not be matched due to lack of partners
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_beginners_list, num=sql_num_con)
        ordered_cities = curs.execute(query).fetchall()
        for iteration in range(len(ordered_cities)*max_guaranteed_beginners):
            query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_beginners_list, num=sql_num_con)
            ordered_cities = curs.execute(query).fetchall()
            for city in ordered_cities:
                number_of_city_beginners = city[city_dict[sql_num_con]]
                max_number_of_city_beginners = city[city_dict[sql_max_con]]
                city = city[city_dict[sql_city]]
                if (max_number_of_city_beginners - number_of_city_beginners) > 0:
                    query = 'SELECT * FROM {tn1} WHERE ({ballroom_level} = ? OR {latin_level} = ?) AND {city} = ?' \
                        .format(tn1=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                                city=sql_city)
                    city_beginners = curs.execute(query, (levels['beg'], levels['beg'], city,)).fetchall()
                    number_of_city_beginners = len(city_beginners)
                    if number_of_city_beginners > 0:
                        random_order = random.sample(range(0, number_of_city_beginners), number_of_city_beginners)
                        beg = city_beginners[random_order[0]]
                        beginner_id = beg[gen_dict[sql_id]]
                        partner_id = find_partner(beginner_id, connection=conn, cursor=curs)
                        create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                        move_selected_contestant(beginner_id, connection=conn, cursor=curs)
                        move_selected_contestant(partner_id, connection=conn, cursor=curs)
                        update_city_beginners(competing_cities, connection=conn, cursor=curs)
        reset_selection_tables(connection=conn, cursor=curs)

    ####################################################################################################################
    # Select guaranteed lions contestants
    ####################################################################################################################
    status_print('')
    status_print('Selecting guaranteed lions for each team...')
    create_city_lions_list(competing_cities, connection=conn, cursor=curs)
    query = 'SELECT sum({max_lion})-sum({num}) FROM {tn}' \
        .format(tn=fixed_lions_list, max_lion=sql_max_con, num=sql_num_con)
    max_iterations = curs.execute(query).fetchone()[0]
    for iteration in range(max_iterations):
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_lions_list, num=sql_num_con)
        ordered_cities = curs.execute(query).fetchall()
        selected_city = None
        for city in ordered_cities:
            if selected_city is None:
                number_of_selected_city_lions = city[city_dict[sql_num_con]]
                max_number_of_city_lions = city[city_dict[sql_max_con]]
                if(max_number_of_city_lions - number_of_selected_city_lions) > 0:
                    selected_city = city[city_dict[sql_city]]
        if selected_city is not None:
            query = get_lions_query()
            selected_city_lions = curs.execute(query, (selected_city,)).fetchall()
            number_of_city_lions = len(selected_city_lions)
            if number_of_city_lions > 0:
                random_order = random.sample(range(0, number_of_city_lions), number_of_city_lions)
                partner_id = None
                for order_city in ordered_cities:
                    if partner_id is None:
                        order_city = order_city[city_dict[sql_city]]
                        if order_city != selected_city:
                            query = get_lions_query()
                            order_city_lions = curs.execute(query, (order_city,)).fetchall()
                            number_of_available_lions = len(order_city_lions)
                            if number_of_available_lions > 0:
                                for num in random_order:
                                    if partner_id is None:
                                        lion = selected_city_lions[num]
                                        lion_id = lion[gen_dict[sql_id]]
                                        query = ' SELECT * FROM {tn} WHERE {id} = ?' \
                                            .format(tn=selected_list, id=sql_id)
                                        lion_available = curs.execute(query, (lion_id,)).fetchone()
                                        if lion_available is None:
                                            partner_id = find_partner(lion_id, connection=conn, cursor=curs,
                                                                      city=order_city)
                                            if partner_id is not None:
                                                create_pair(lion_id, partner_id, connection=conn, cursor=curs)
                                                move_selected_contestant(lion_id, connection=conn, cursor=curs)
                                                move_selected_contestant(partner_id, connection=conn, cursor=curs)
                                                update_city_lions(competing_cities, connection=conn, cursor=curs)
    reset_selection_tables(connection=conn, cursor=curs)

    ####################################################################################################################
    # Select remaining contestants
    ####################################################################################################################
    status_print('')
    status_print('Selecting the bulk of contestants that have signed up...')
    select_bulk(limit=max_contestants-buffer_for_selection, connection=conn, cursor=curs)
    status_print('')
    status_print("--- Done in %.3f seconds ---" % (time.time() - start_time))
    status_print('')

    ####################################################################################################################
    # Collect user data from main selection
    ####################################################################################################################
    collect_city_overview(source_table=fixed_beginners_list, target_table=beginners_list, users=levels['beg'],
                          cursor=curs, connection=conn, collect_data=True)
    collect_city_overview(source_table=fixed_lions_list, target_table=lions_list, users=lions,
                          cursor=curs, connection=conn, collect_data=True)
    collect_city_overview(source_table=selected_list, target_table=contestants_list, users=contestants,
                          cursor=curs, connection=conn, collect_data=True)
    if gather_stats:
        query = 'SELECT * FROM {tn}'.format(tn=signup_list)
        all_dancers = curs.execute(query).fetchall()
        query = 'SELECT name FROM sqlite_master WHERE type = ? AND name = ?'
        individual_table_exists = len(curs.execute(query, ('table', individual_list)).fetchall())
        if individual_table_exists == 0:
            query = 'CREATE TABLE IF NOT EXISTS {tn} ({run} INTEGER PRIMARY KEY, '\
                .format(tn=individual_list, run=sql_run)
            for dancer in all_dancers:
                dancer_id = dancer[gen_dict[sql_id]]
                query += ('"' + str(dancer_id) + '" INT, ')
            query = query[:-2]
            query += ')'
            curs.execute(query)
        query = 'SELECT {run} FROM {tn}'.format(tn=individual_list, run=sql_run)
        this_run = len(curs.execute(query).fetchall()) + 1
        query = 'INSERT INTO {tn} ({run}) VALUES (?)'.format(tn=individual_list, run=sql_run)
        curs.execute(query, (this_run,))
        for dancer in all_dancers:
            dancer_id = dancer[gen_dict[sql_id]]
            query = 'SELECT * FROM {tn} WHERE {id} = ?'.format(id=sql_id, tn=selected_list)
            dancer_selected = len(curs.execute(query, (dancer_id,)).fetchall())
            query = 'UPDATE {tn} SET "{col}" = ? WHERE {run} = ?'.format(tn=individual_list, run=sql_run, col=dancer_id)
            curs.execute(query, (dancer_selected, this_run))
        conn.commit()
    # Update status
    status_update()
    # Close cursor and connection
    curs.close()
    conn.close()
    # Enable the CLI
    cli_text.config(state=NORMAL)


if __name__ == "__main__":
    root = Tk()
    root.geometry("1600x900")
    root.state('zoomed')
    root.title('NTDS 2018 Selection (BETA)')
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
    cli_text = Entry(master=frame, width=int(status_text_width*27/20))
    cli_text.grid(row=2, column=0, padx=pad_in, pady=pad_out)
    cli_text.bind('<Return> ', cli_parser)
    data_help_frame = Frame(master=frame)
    data_help_frame.grid(row=0, column=2, rowspan=4, columnspan=3)
    data_text = Text(master=data_help_frame, width=total_width-status_text_width, height=50,
                     padx=pad_in, pady=pad_in, wrap=WORD, state=DISABLED)
    data_text.grid(row=0, column=0, padx=pad_out, columnspan=3)
    padding_frame = Frame(master=data_help_frame, height=16)
    padding_frame.grid(row=1, column=0)
    start_button = Button(master=data_help_frame, text='Start new selection database', command=main_selection)
    start_button.grid(row=2, column=0, padx=pad_out, pady=pad_in)
    update_button = Button(master=data_help_frame, text='Select existing database', command=select_database)
    update_button.grid(row=2, column=1, padx=pad_out)
    options_button = Button(master=data_help_frame, text='Settings', command=options_menu)
    options_button.grid(row=2, column=2, padx=pad_out)
    welcome_text()
    print_ntds_config()
    cli_text.focus_set()
    select_db = EntryBox
    select_db.root = root
    root.mainloop()
