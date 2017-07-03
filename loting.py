# Main file for the program
import sqlite3 as sql
import openpyxl
# from openpyxl import Workbook
from random import randint
import random
import time
from tabulate import tabulate
# import logging
# logging.basicConfig(filename='loting.log', filemode='w', level=logging.DEBUG)
# logger = logging.getLogger()
# logger.setLevel(logging.DEBUG)
# logger.addHandler(logging.StreamHandler())
from tkinter import *

debug = True

# TODO CLI
# TODO controle programma voor inschrijflijst

# init stuff
participating_teams = 'deelnemende_teams.xlsx'
template = 'NTDS_Template.xlsx'
selection = 'NTDS_Selection.xlsx'

# boundaries
max_contestants = 400
max_fixed_beginner_pairs = 2
max_fixed_beginners = max_fixed_beginner_pairs*2
beginner_signup_cutoff = 80
max_fixed_lion_pairs = 5
max_fixed_lion_contestants = max_fixed_lion_pairs*2
buffer_for_selection = 40

# Names
database_name = 'main_data.db'
selected_database = database_name
signup_list = 'signup_list'
selection_list = 'selection_list'
selected_list = 'selected_list'
team_list = 'team_list'
partners_list = 'partners_list'
ref_partner_list = 'reference_partner_list'
fixed_beginners_list = 'fixed_beginners'
fixed_lions_list = 'fixed_lions'
beginners_list = 'beginners'
lions_list = 'lions'
contestants_list = 'contestants'
individual_list = 'individuals'
sql_run = 'run_number'
Beginners = 'Beginners'
Lions = 'Lions'
contestants = 'contestants'

# more names
breiten = 'Breiten'
beginner = 'Beginner'
open_class = 'Open'
lead = 'Lead'
follow = 'Follow'

# options
yes = 'Ja'
maybe = 'Misschien'

# selectable options
NTDS_ballroom_level = {0: beginner, 1: breiten, 2: open_class}
NTDS_latin_level = {0: beginner, 1: breiten, 2: open_class}
NTDS_ballroom_role = {0: lead, 1: follow, 2: ''}
NTDS_latin_role = {0: lead, 1: follow, 2: ''}
NTDS_ballroom_mandatory_blind_date = {'yes': yes, 0: ''}
NTDS_latin_mandatory_blind_date = {'yes': yes, 0: ''}
NTDS_team_captain = {'yes': yes, 0: ''}
NTDS_first_aid = {'yes': yes, 'maybe': maybe, 0: ''}
NTDS_emergency_response_officer = {'yes': yes, 'maybe': maybe, 0: ''}
NTDS_ballroom_jury = {'yes': yes, 'maybe': maybe, 0: ''}
NTDS_latin_jury = {'yes': yes, 'maybe': maybe, 0: ''}
NTDS_student = {'yes': yes, 0: ''}
NTDS_sleeping_location = {'yes': yes, 0: ''}
NTDS_current_volunteer = {'yes': yes, 'maybe': maybe, 0: ''}
NTDS_past_volunteer = {'yes': yes, 0: ''}
NTDS_options = {'ballroom_level': NTDS_ballroom_level, 'latin_level': NTDS_latin_level,
                'ballroom_role': NTDS_ballroom_role, 'latin_role': NTDS_latin_role,
                'ballroom_mandatory_blind_date': NTDS_ballroom_mandatory_blind_date,
                'latin_mandatory_blind_date': NTDS_latin_mandatory_blind_date,
                'team_captain': NTDS_team_captain,
                'first_aid': NTDS_first_aid, 'emergency_response_officer': NTDS_emergency_response_officer,
                'ballroom_jury': NTDS_ballroom_jury, 'latin_jury': NTDS_latin_jury,
                'student': NTDS_student, 'sleeping_location': NTDS_sleeping_location,
                'current_volunteer': NTDS_current_volunteer, 'past_volunteer': NTDS_past_volunteer}

# Participants Lion points
lion_participants = [beginner, breiten]

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
partner_dict = {sql_lead: 1, sql_follow: 2, sql_city_lead: 3, sql_city_follow: 4,
                sql_ballroom_level_lead: 5, sql_ballroom_level_follow: 6,
                sql_latin_level_lead: 7, sql_latin_level_follow: 8}

# SQL Table column names and dictionary for city list
sql_num_con = 'number_of_contestants'
sql_max_con = 'max_contestants'
city_dict = {sql_city: 0, sql_num_con: 1, sql_max_con: 2}

# General query formats
drop_table_query = 'DROP TABLE IF EXISTS {};'
dancers_list_query = 'CREATE TABLE {table_name} ({id} INT PRIMARY KEY, ' \
                     '{first_name} TEXT, {ln_prefix} TEXT, {last_name} TEXT, {email} TEXT, ' \
                     '{ballroom_level} TEXT, {latin_level} TEXT, {ballroom_partner} INT, {latin_partner} INT, ' \
                     '{ballroom_role} TEXT, {latin_role} TEXT, ' \
                     '{ballroom_mandatory_blind_date} TEXT, {latin_mandatory_blind_date} TEXT, {team_captain} TEXT,' \
                     '{first_aid} TEXT, {emergency_response_officer} TEXT, {ballroom_jury} TEXT, {latin_jury} TEXT, ' \
                     '{student} TEXT, {sleeping_location} TEXT, {diet} TEXT, ' \
                     '{current_volunteer} TEXT, {previous_volunteer} TEXT,' \
                     '{city} TEXT);'\
                     .format(table_name={}, id=sql_id, first_name=sql_first_name, ln_prefix=sql_ln_prefix,
                             last_name=sql_last_name, email=sql_email,
                             ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
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
                    '{latin_level_lead} TEXT, {latin_level_follow} TEXT);' \
    .format(tn={}, no=sql_no, lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow,
            ballroom_level_lead=sql_ballroom_level_lead, ballroom_level_follow=sql_ballroom_level_follow,
            latin_level_lead=sql_latin_level_lead, latin_level_follow=sql_latin_level_follow)
city_list_query = 'CREATE TABLE {tn} ({city} TEXT, {beg} INT, {max_beg} INT);' \
    .format(tn={}, city=sql_city, beg=sql_num_con, max_beg=sql_max_con)


def find_partner(identifier, connection, cursor, city=None):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    status_print('Attempting to find a partner for dancer {id}'.format(id=identifier))
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
        if partner_id is None:
            potential_partners = cursor.execute(query, (ballroom_level, latin_level, ballroom_role, latin_role, team))\
                .fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {latin_level} = ? AND {ballroom_partner} = "" AND ' \
                '{latin_partner} = "" AND {ballroom_role} != ? AND latin_role != ? '\
            .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level,
                    ballroom_partner=sql_ballroom_partner, latin_partner=sql_latin_partner,
                    ballroom_role=sql_ballroom_role, latin_role=sql_latin_role)
        if city is None:
            query += ' AND {team} != ?'.format(team=sql_city)
        else:
            query += ' AND {team} = ?'.format(team=sql_city)
            team = city
        # Try to find a partner for a beginner, beginner combination
        if all([ballroom_level == beginner, latin_level == beginner,
                partner_id is None, number_of_potential_partners == 0]):
            potential_partners += cursor.execute(query, (beginner, '', ballroom_role, '', team)).fetchall()
            potential_partners += cursor.execute(query, ('', beginner, '', latin_role, team)).fetchall()
        # Try to find a partner for a beginner, Null combination
        if all([ballroom_level == beginner, latin_level == '',
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (beginner, beginner, ballroom_role, ballroom_role, team))\
                .fetchall()
        # Try to find a partner for a Null, beginner combination
        if all([ballroom_level == '', latin_level == beginner,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (beginner, beginner, latin_role, latin_role, team))\
                .fetchall()
        # Try to find a partner for a breiten, breiten combination
        if all([ballroom_level == breiten, latin_level == breiten,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, '', ballroom_role, '', team)).fetchall()
            potential_partners += cursor.execute(query, ('', breiten, '', latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, (breiten, open_class, ballroom_role, latin_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (open_class, breiten, ballroom_role, latin_role, team))\
                .fetchall()
        # Try to find a partner for a breiten, Null combination
        if all([ballroom_level == breiten, latin_level == '',
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, breiten, ballroom_role, ballroom_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (breiten, open_class, ballroom_role, ballroom_role, team))\
                .fetchall()
        # Try to find a partner for a Null, breiten combination
        if all([ballroom_level == '', latin_level == breiten,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, breiten, latin_role, latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, (open_class, breiten, latin_role, latin_role, team)).fetchall()
        # Try to find a partner for a breiten, Open combination
        if all([ballroom_level == breiten, latin_level == open_class,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, breiten, ballroom_role, latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, (breiten, '', ballroom_role, '', team)).fetchall()
            potential_partners += cursor.execute(query, (open_class, open_class, ballroom_role, latin_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, ('', open_class, '', latin_role, team)).fetchall()
        # Try to find a partner for a Open, Breiten combination
        if all([ballroom_level == open_class, latin_level == breiten,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, breiten, ballroom_role, latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, ('', breiten, '', latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, (open_class, open_class, ballroom_role, latin_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (open_class, '', ballroom_role, '', team)).fetchall()
        # Try to find a partner for a Open, Open combination
        if all([ballroom_level == open_class, latin_level == open_class,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, open_class, ballroom_role, latin_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (open_class, breiten, ballroom_role, latin_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (open_class, '', ballroom_role, '', team)).fetchall()
            potential_partners += cursor.execute(query, ('', open_class, '', latin_role, team)).fetchall()
        # Try to find a partner for a Open, Null combination
        if all([ballroom_level == open_class, latin_level == '',
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (open_class, breiten, ballroom_role, ballroom_role, team))\
                .fetchall()
            potential_partners += cursor.execute(query, (open_class, open_class, ballroom_role, ballroom_role, team))\
                .fetchall()
        # Try to find a partner for a Null, Open combination
        if all([ballroom_level == '', latin_level == open_class,
                number_of_potential_partners == 0, partner_id is None]):
            potential_partners += cursor.execute(query, (breiten, open_class, latin_role, latin_role, team)).fetchall()
            potential_partners += cursor.execute(query, (open_class, open_class, latin_role, latin_role, team))\
                .fetchall()
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
    if first_dancer_ballroom_role == follow or first_dancer_latin_role == follow:
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
            '{ballroom_level_lead}, {ballroom_level_follow}, {latin_level_lead}, latin_level_follow) ' \
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?)'\
        .format(tn=partners_list,
                lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow,
                ballroom_level_lead=sql_ballroom_level_lead, ballroom_level_follow=sql_ballroom_level_follow,
                latin_level_lead=sql_latin_level_lead, latin_level_follow=sql_latin_level_follow)
    cursor.execute(query, (first_dancer_id, second_dancer_id, first_dancer_team, second_dancer_team,
                           first_dancer_ballroom_level, second_dancer_ballroom_level,
                           first_dancer_latin_level, second_dancer_latin_level))
    query = 'INSERT INTO {tn} ({lead}, {follow}, {lead_city}, {follow_city}, ' \
            '{ballroom_level_lead}, {ballroom_level_follow}, {latin_level_lead}, latin_level_follow) ' \
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


def create_city_beginners_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn1} WHERE ({ballroom_level} = ? OR {latin_level} = ?) AND {team} = ?' \
            .format(tn1=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level, team=sql_city)
        max_city_beginners = len(cursor.execute(query, (beginner, beginner, city)).fetchall())
        if max_city_beginners > max_fixed_beginners:
            max_city_beginners = max_fixed_beginners
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=fixed_beginners_list)
        cursor.execute(query, (city, 0, max_city_beginners))
    connection.commit()


def select_bulk(limit, connection, cursor):
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
                    if partner_id is not None:
                        create_pair(dancer_id, partner_id, connection=connection, cursor=cursor)
                        move_selected_contestant(dancer_id, connection=connection, cursor=cursor)
                        move_selected_contestant(partner_id, connection=connection, cursor=cursor)
            query = 'SELECT * FROM {tn}'.format(tn=selected_list)
            remaining_dancers = cursor.execute(query).fetchall()
            number_of_selected_dancers = len(remaining_dancers)
            if number_of_selected_dancers >= limit:
                break
    connection.commit()
    reset_selection_tables(connection=connection, cursor=cursor)


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
        query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ?' \
            .format(tn=partners_list, city_lead=sql_city_lead)
        number_of_city_lions = len(cursor.execute(query, (city,)).fetchall())
        query = 'SELECT * FROM {tn} WHERE {city_follow} LIKE ?' \
            .format(tn=partners_list, city_follow=sql_city_follow)
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
    tables_to_drop = [signup_list, selection_list, selected_list, team_list, partners_list, ref_partner_list,
                      fixed_beginners_list, fixed_lions_list]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        cursor.execute(query)
    # Create new tables
    dancer_list_tables = [signup_list, selection_list, selected_list]
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
    # Create Workbook object and a Worksheet from it
    wb = openpyxl.load_workbook(participating_teams)
    ws = wb.worksheets[0]
    max_row = max_rc('row', ws)
    # Empty 2d list that will contain the signup sheet filenames for each of the teams
    competing_cities_array = []
    # Fill up list with competing cities
    for r in range(2, max_row + 1):
        competing_cities_array \
            .append([ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value, ws.cell(row=r, column=3).value])
    query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=team_list)
    for row in competing_cities_array:
        cursor.execute(query, row)
    connection.commit()
    return competing_cities_array


def reset_selection_tables(connection, cursor):
    """"Temp"""
    query = drop_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()
    query = paren_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()


def collect_city_overview(source_table, target_table, users, cursor, connection):
    """"Temp"""
    if source_table == selected_list:
        query = 'SELECT {city}, COUNT() FROM {tn} GROUP BY {city}'.format(tn=source_table, city=sql_city)
    else:
        query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=source_table, city=sql_city)
    ordered_cities = cursor.execute(query).fetchall()
    if debug:
        query = 'CREATE TABLE IF NOT EXISTS {tn} ' \
                '(id INTEGER PRIMARY KEY AUTOINCREMENT, ' \
                '{c1} INT, {c2} INT, {c3} INT, {c4} INT, {c5} INT, {c6} INT, {c7} INT, {c8} INT, {c9} INT, ' \
                '{c10} INT, {c11} INT)' \
            .format(tn=target_table, c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0],
                    c4=ordered_cities[3][0], c5=ordered_cities[4][0], c6=ordered_cities[5][0], c7=ordered_cities[6][0],
                    c8=ordered_cities[7][0], c9=ordered_cities[8][0], c10=ordered_cities[9][0],
                    c11=ordered_cities[10][0])
        cursor.execute(query)
        query = 'INSERT INTO {tn} ({c1},{c2},{c3},{c4},{c5},{c6},{c7},{c8},{c9},{c10},{c11}) ' \
                'VALUES (?,?,?,?,?,?,?,?,?,?,?)' \
            .format(tn=target_table, c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0],
                    c4=ordered_cities[3][0], c5=ordered_cities[4][0], c6=ordered_cities[5][0], c7=ordered_cities[6][0],
                    c8=ordered_cities[7][0], c9=ordered_cities[8][0], c10=ordered_cities[9][0],
                    c11=ordered_cities[10][0])
        cursor.execute(query, (ordered_cities[0][1], ordered_cities[1][1], ordered_cities[2][1], ordered_cities[3][1],
                               ordered_cities[4][1], ordered_cities[5][1], ordered_cities[6][1], ordered_cities[7][1],
                               ordered_cities[8][1], ordered_cities[9][1], ordered_cities[10][1]))
    for city in ordered_cities:
        overview = 'Number of selected {users} from {city} is: {number}'\
            .format(city=city[0], number=city[1], users=users)
        status_print(overview)
    status_print('')
    connection.commit()


def print_ntds_config():
    """"Temp"""
    status_print('Contestant numbers for this NTDS selection:')
    status_print('Maximum number of contestants: {num}'.format(num=max_contestants))
    status_print('Guaranteed beginners per team: {num}'.format(num=max_fixed_beginners))
    status_print('Guaranteed Lion contestants per team: {num}'.format(num=max_fixed_lion_contestants))
    status_print('Cutoff for selecting all beginners: {num}'.format(num=beginner_signup_cutoff))
    status_print('Buffer for selecting contestants at the end: {num}'.format(num=buffer_for_selection))
    status_print('Levels participating for the Lion:')
    for level in lion_participants:
        status_print('\t{lvl}'.format(lvl=level))


def welcome_text():
    """"Text displayed when opening the program for the first time."""
    status_print('Welcome to the NTDS Selection!')
    status_print('')
    status_print('You can start a new selection, update an existing selection, '
                 'or change the options with the buttons in the bottom right corner.')
    status_print('')
    data_text.config(state=NORMAL)
    data_text.delete('1.0', END)
    data_text.insert(END, 'Data about the number of contestants, First Aid Officers, required sleeping locations, etc. '
                          'will be displayed here once a database has been selected.')
    data_text.config(state=DISABLED)
    help_text.config(state=NORMAL)
    help_text.insert(END, 'Some helpful commands are:\n')
    help_text.insert(END, 'list_available:\n')
    help_text.insert(END, 'Lists all dancers available for selection\n')
    help_text.insert(END, 'list_level: {level=beginners/breiten/open}\n')
    help_text.insert(END, 'Lists all dancers of the given level available for selection\n')
    help_text.insert(END, 'list_fa:\n')
    help_text.insert(END, 'Lists all available First Aid Officers\n')
    help_text.insert(END, 'list_ero:\n')
    help_text.insert(END, 'Lists all available Emergency Response Officers\n')
    help_text.insert(END, '-add n:\n')
    help_text.insert(END, 'Selects contestant number "n" (and a potential partner) for the NTDS\n')
    help_text.config(state=DISABLED)

db_commands = ['list_fa', 'list_ero', 'list_available', 'list_beginners', 'list_breiten', 'list_open',
                   'print_contestants',
                   'finish_selection', 'export']
def command_help_text():
    """"Temp"""
    status_print('Listing all commands...')
    status_print('list_fa:')
    status_print('Lists all dancers available for selection that are a qualified First Aid Officer')
    status_print('list_ero:')
    status_print('Lists all dancers available for selection that are a qualified Emergency Response Officer')
    status_print('list_available:')
    status_print('Lists all dancers available for selection')
    status_print('list_beginners:')
    status_print('Lists all Beginners available for selection')
    status_print('list_breiten:')
    status_print('Lists all dancers with at least one level Breitensport available for selection')
    status_print('list_open:')
    status_print('Lists all dancers with at least one level Open Class available for selection')
    status_print('print_contestants')
    status_print('Prints a list of how much contestants each city has had selected')
    status_print('finish_selection')
    status_print('Finishes the NTDS selection by adding random contestants')


def status_print(status_message):
    """"Temp"""
    status_text.config(state=NORMAL)
    status_text.insert(END, status_message)
    status_text.insert(END, '\n')
    status_text.update()
    status_text.see(END)
    status_text.config(state=DISABLED)


def print_table(table):
    """"Temp"""
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
                            dancer[gen_dict[sql_current_volunteer]], dancer[gen_dict[sql_previous_volunteer]]]
        formatted_table.append(formatted_dancer)
    # print_table_header = ['Id', 'Name', 'Ballroom level', 'Latin level', 'Ballroom partner', 'Latin partner',
    #                       'Ballroom role', 'Latin role', \
    #                       'Ballroom mandatory blind date', 'Latin mandatory blind date',
    #                       'First Aid', 'Emergency Response Officer', 'Ballroom jury', 'Latin jury', 'Student',
    #                       'Sleeping location', 'Volunteer', 'Past volunteer']
    print_table_header = ['Id', 'Name', 'B-lvl', 'L-lvl', 'B-part', 'L-part',
                          'B-role', 'L-role', 'B-date', 'L-date',
                          'F.A.', 'E.R.O.', 'B-jury', 'L-jury', 'Student',
                          'Sleeping', 'Volunteer', 'Past volunteer']
    # table_cutoff = 24
    # while len(formatted_table) > table_cutoff:
    #     status_print(tabulate(formatted_table[:table_cutoff], headers=print_table_header))
    #     status_print('')
    #     formatted_table = formatted_table[table_cutoff:]
    # status_print(tabulate(formatted_table[:table_cutoff], headers=print_table_header))
    status_print(tabulate(formatted_table, headers=print_table_header))
    # status_print('')


def status_update(cursor):
    """"Temp"""
    query = 'SELECT * FROM {tn}'.format(tn=selected_list)
    number_of_contestants = len(cursor.execute(query).fetchall())
    query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? AND {ballroom_role} = ?'\
        .format(tn=selected_list, ballroom_level=sql_ballroom_level, ballroom_role=sql_ballroom_role)
    number_of_beginner_ballroom_leads = len(cursor.execute(query, (beginner, lead)).fetchall())
    number_of_breiten_ballroom_leads = len(cursor.execute(query, (breiten, lead)).fetchall())
    number_of_open_ballroom_leads = len(cursor.execute(query, (open_class, lead)).fetchall())
    number_of_beginner_ballroom_follows = len(cursor.execute(query, (beginner, follow)).fetchall())
    number_of_breiten_ballroom_follows = len(cursor.execute(query, (breiten, follow)).fetchall())
    number_of_open_ballroom_follows = len(cursor.execute(query, (open_class, follow)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {latin_level} = ? AND {latin_role} = ?' \
        .format(tn=selected_list, latin_level=sql_latin_level, latin_role=sql_latin_role)
    number_of_beginner_latin_leads = len(cursor.execute(query, (beginner, lead)).fetchall())
    number_of_breiten_latin_leads = len(cursor.execute(query, (breiten, lead)).fetchall())
    number_of_open_latin_leads = len(cursor.execute(query, (open_class, lead)).fetchall())
    number_of_beginner_latin_follows = len(cursor.execute(query, (beginner, follow)).fetchall())
    number_of_breiten_latin_follows = len(cursor.execute(query, (breiten, follow)).fetchall())
    number_of_open_latin_follows = len(cursor.execute(query, (open_class, follow)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {first_aid} = ?'.format(tn=selected_list, first_aid=sql_first_aid)
    number_of_first_aid_yes = len(cursor.execute(query, (NTDS_options['first_aid']['yes'],)).fetchall())
    number_of_first_aid_maybe = len(cursor.execute(query, (NTDS_options['first_aid']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {first_aid} = ?'\
        .format(tn=selected_list, first_aid=sql_emergency_response_officer)
    number_of_emergency_response_officer_yes = \
        len(cursor.execute(query, (NTDS_options['emergency_response_officer']['yes'],)).fetchall())
    number_of_emergency_response_officer_maybe = \
        len(cursor.execute(query, (NTDS_options['emergency_response_officer']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE ballroom_level = ? AND ballroom_mandatory_blind_date = ?' \
        .format(tn=selected_list, ballroom_level=sql_ballroom_level,
                ballroom_mandatory_blind_date=sql_ballroom_mandatory_blind_date)
    number_of_mandatory_breiten_ballroom_blind_daters = len(cursor.execute(query, (breiten, yes)).fetchall())
    query = 'SELECT * FROM {tn} WHERE latin_level = ? AND latin_mandatory_blind_date = ?' \
        .format(tn=selected_list, latin_level=sql_latin_level,
                latin_mandatory_blind_date=sql_latin_mandatory_blind_date)
    number_of_mandatory_breiten_latin_blind_daters = len(cursor.execute(query, (breiten, yes)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {ballroom_jury} = ?'.format(tn=selected_list, ballroom_jury=sql_ballroom_jury)
    number_of_ballroom_jury_yes = len(cursor.execute(query, (NTDS_options['ballroom_jury']['yes'],)).fetchall())
    number_of_ballroom_jury_maybe = len(cursor.execute(query, (NTDS_options['ballroom_jury']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {latin_jury} = ?'.format(tn=selected_list, latin_jury=sql_latin_jury)
    number_of_latin_jury_yes = len(cursor.execute(query, (NTDS_options['latin_jury']['yes'],)).fetchall())
    number_of_latin_jury_maybe = len(cursor.execute(query, (NTDS_options['latin_jury']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {current_volunteer} = ?'\
        .format(tn=selected_list, current_volunteer=sql_current_volunteer)
    number_of_current_volunteer_yes = \
        len(cursor.execute(query, (NTDS_options['current_volunteer']['yes'],)).fetchall())
    number_of_current_volunteer_maybe = \
        len(cursor.execute(query, (NTDS_options['current_volunteer']['maybe'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {past_volunteer} = ?'\
        .format(tn=selected_list, past_volunteer=sql_previous_volunteer)
    number_of_past_volunteer = len(cursor.execute(query, (NTDS_options['past_volunteer']['yes'],)).fetchall())
    query = 'SELECT * FROM {tn} WHERE {sleeping_location} = ?' \
        .format(tn=selected_list, sleeping_location=sql_sleeping_location)
    number_of_sleeping_spots = len(cursor.execute(query, (NTDS_options['sleeping_location']['yes'],)).fetchall())
    data_text.config(state=NORMAL)
    data_text.delete('1.0', END)
    data_text.insert(END, 'Total number of contestants: {num}\n'.format(num=number_of_contestants))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Beginners Ballroom Leads: {num}\n'.format(num=number_of_beginner_ballroom_leads))
    data_text.insert(END, 'Beginners Ballroom Follows: {num}\n'.format(num=number_of_beginner_ballroom_follows))
    data_text.insert(END, 'Beginners Latin Leads: {num}\n'.format(num=number_of_beginner_latin_leads))
    data_text.insert(END, 'Beginners Latin Follows: {num}\n'.format(num=number_of_beginner_latin_follows))
    data_text.insert(END, '\n')
    data_text.insert(END, 'Breitensport Ballroom Leads: {num}\n'.format(num=number_of_breiten_ballroom_leads))
    data_text.insert(END, 'Breitensport Ballroom Follows: {num}\n'.format(num=number_of_breiten_ballroom_follows))
    data_text.insert(END, 'Breitensport Latin Leads: {num}\n'.format(num=number_of_breiten_latin_leads))
    data_text.insert(END, 'Breitensport Latin Follows: {num}\n'.format(num=number_of_breiten_latin_follows))
    data_text.insert(END, '\n')
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


def cli_parser(*args):
    """"Temp"""
    wip = 'Work in progress...'
    command = cli_text.get()
    cli_text.delete(0, END)
    connection = sql.connect(selected_database)
    cli_curs = connection.cursor()
    open_commands = ['echo', 'help']
    db_commands = ['list_fa', 'list_ero', 'list_available', 'list_beginners', 'list_breiten', 'list_open',
                   'print_contestants',
                   'finish_selection', 'export']
    if (command in db_commands or command.startswith('-')) and selected_database != '':
        if command.startswith('-add '):
            command = command[5:]
            selected_id = int(command)
            partner_id = find_partner(selected_id, connection=connection, cursor=cli_curs)
            create_pair(selected_id, partner_id, connection=connection, cursor=cli_curs)
            move_selected_contestant(selected_id, connection=connection, cursor=cli_curs)
            move_selected_contestant(partner_id, connection=connection, cursor=cli_curs)
        elif command == 'list_fa':
            query = 'SELECT * FROM {tn} WHERE {first_aid} = ?'.format(tn=selection_list, first_aid=sql_first_aid)
            available_contestants = cli_curs.execute(query, (NTDS_options['first_aid']['maybe'],)).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('Contestants that MIGHT want to volunteer as a First Aid Officer:')
                status_print('')
                print_table(available_contestants)
                status_print('')
            available_contestants = cli_curs.execute(query, (NTDS_options['first_aid']['yes'],)).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('Contestants that want to volunteer as a First Aid Officer:')
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no volunteers available for selection that are a qualified First Aid Officer')
                status_print('')
        elif command == 'list_ero':
            query = 'SELECT * FROM {tn} WHERE {ero} = ?'.format(tn=selection_list, ero=sql_emergency_response_officer)
            available_contestants = cli_curs.execute(query, (NTDS_options['emergency_response_officer']['maybe'],))\
                .fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('Contestants that MIGHT want to volunteer as a Emergency Response Officer:')
                status_print('')
                print_table(available_contestants)
                status_print('')
            available_contestants = cli_curs.execute(query, (NTDS_options['emergency_response_officer']['yes'],)) \
                .fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('Contestants that want to volunteer as a Emergency Response Officer:')
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no volunteers available for selection that are a qualified '
                             'Emergency Response Officer')
                status_print('')
        elif command == 'list_available':
            query = 'SELECT * FROM {tn}'.format(tn=selection_list)
            available_contestants = cli_curs.execute(query).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('All {num} contestants that are available for selection for the NTDS:'
                             .format(num=len(available_contestants)))
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no contestants available for selection for the NTDS')
                status_print('')
        elif command == 'list_beginners':
            query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? or {latin_level} = ?'\
                .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level)
            available_contestants = cli_curs.execute(query, (beginner, beginner)).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('All {num} {lvl} dancers that are available for selection for the NTDS:'
                             .format(num=len(available_contestants), lvl=beginner))
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no {lvl} dancers available for selection for the NTDS'.format(lvl=beginner))
                status_print('')
        elif command == 'list_breiten':
            query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? or {latin_level} = ?' \
                .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level)
            available_contestants = cli_curs.execute(query, (breiten, breiten)).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('All {num} {lvl} dancers that are available for selection for the NTDS:'
                             .format(num=len(available_contestants), lvl=breiten))
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no {lvl} dancers available for selection for the NTDS'.format(lvl=breiten))
                status_print('')
        elif command == 'list_open':
            query = 'SELECT * FROM {tn} WHERE {ballroom_level} = ? or {latin_level} = ?' \
                .format(tn=selection_list, ballroom_level=sql_ballroom_level, latin_level=sql_latin_level)
            available_contestants = cli_curs. execute(query, (open_class, open_class)).fetchall()
            if len(available_contestants) > 0:
                status_print('')
                status_print('All {num} {lvl} dancers that are available for selection for the NTDS:'
                             .format(num=len(available_contestants), lvl=open_class))
                status_print('')
                print_table(available_contestants)
                status_print('')
            else:
                status_print('There are no {lvl} dancers available for selection for the NTDS'.format(lvl=open_class))
                status_print('')
        elif command == 'finish_selection':
            select_bulk(max_contestants, connection=connection, cursor=cli_curs)
        elif command == 'print_contestants':
            status_print('')
            collect_city_overview(source_table=selected_list, target_table=contestants_list, users=contestants,
                                  cursor=cli_curs, connection=connection)
        elif command == 'export':
            # otp = selection.replace(".", ("_" + str(time.time()) + "."))
            # wb = openpyxl.Workbook()
            # wb.save(otp)
            status_print(wip)
    elif command in open_commands:
        if command == 'echo':
            status_print(command)
        elif command == 'help':
            command_help_text()
    else:
        if selected_database == '':
            status_print('Command not available, no open database')
        else:
            status_print('Unknown command: "{command}"'.format(command=command))
    if selected_database != '':
        status_update(cursor=cli_curs)
    cli_curs.close()
    connection.close()


def main_selection():
    ####################################################################################################################
    # Extract data from sign-up sheets
    ####################################################################################################################
    start_time = time.time()
    timestamp = start_time
    # Connect to database and create a cursor
    conn = sql.connect(selected_database)
    curs = conn.cursor()
    # Create SQL tables
    create_tables(connection=conn, cursor=curs)
    # Create competing cities list
    competing_teams = create_competing_teams(connection=conn, cursor=curs)
    competing_cities = [row[team_dict[sql_city]] for row in competing_teams]
    # Get maximum number of columns
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    max_col = max_rc('col', ws)
    # Copy the signup list from every team into the SQL database
    total_signup_list = list()
    total_number_of_contestants = 0
    for team in competing_teams:
        city = team[team_dict[sql_city]]
        team_signup_list = team[team_dict[sql_signup_list]]
        # Get maximum number of rows and extract signup list
        wb = openpyxl.load_workbook(team_signup_list)
        ws = wb.worksheets[0]
        max_row = max_rc('row', ws)
        city_signup_list = list(ws.iter_rows(min_col=1, min_row=2, max_col=max_col, max_row=max_row))
        # Convert data to 2d list, replace None values with an empty string,
        # increase the id numbers so that there are no duplicates, and add the city to the contestant
        city_signup_list = [[cell.value for cell in row] for row in city_signup_list]
        city_signup_list = [['' if elem is None else elem for elem in row] for row in city_signup_list]
        city_signup_list = [[elem+total_number_of_contestants if isinstance(elem, int) else elem for elem in row]
                            for row in city_signup_list]
        for row in city_signup_list:
            row.append(city)
        total_signup_list.extend(city_signup_list)
        # Copy the contestants to the SQL database
        query = 'INSERT INTO {} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);'.format(signup_list)
        for row in city_signup_list:
            curs.execute(query, row)
        query = 'INSERT INTO {} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);'.format(selection_list)
        for row in city_signup_list:
            curs.execute(query, row)
        total_number_of_contestants += max_row-1
    conn.commit()

    ####################################################################################################################
    # Select the team captains and (virtual) partners
    ####################################################################################################################
    status_print('Selecting team captains...')
    query = 'SELECT * FROM {tn1} WHERE {tc} = "Ja"'.format(tn1=selection_list, tc=sql_team_captain)
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
    all_beginners = curs.execute(query, (beginner, beginner)).fetchall()
    number_of_signed_beginners = len(all_beginners)
    if number_of_signed_beginners <= beginner_signup_cutoff:
        status_print('Less than {num} Beginners signed up.'.format(num=beginner_signup_cutoff+1))
        status_print('Matching up as much couples as possible and selecting everyone...')
        for beg in all_beginners:
            beg_id = beg[gen_dict[sql_id]]
            query = 'SELECT * FROM {tn1} WHERE {id} = ?'.format(tn1=selected_list, id=sql_id)
            beginner_selected = curs.execute(query, (beg_id,)).fetchone()
            if beginner_selected is None:
                partner_id = find_partner(beg_id, connection=conn, cursor=curs)
                create_pair(beg_id, partner_id, connection=conn, cursor=curs)
                move_selected_contestant(beg_id, connection=conn, cursor=curs)
                move_selected_contestant(partner_id, connection=conn, cursor=curs)
        conn.commit()

    ####################################################################################################################
    # Select beginners if more people have signed than the given cutoff
    ####################################################################################################################
    if number_of_signed_beginners > beginner_signup_cutoff:
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
                selected_city_beginners = curs.execute(query, (beginner, beginner, selected_city,)).fetchall()
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
                                order_city_beginners = curs.execute(query, (beginner, beginner, order_city,)).fetchall()
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
        for iteration in range(len(ordered_cities)*max_fixed_beginners):
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
                    city_beginners = curs.execute(query, (beginner, beginner, city,)).fetchall()
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
    status_print('Selecting the bulk of contestants that have signed up...')
    query = 'SELECT * FROM {tn}'.format(tn=selection_list)
    available_dancers = curs.execute(query).fetchall()
    number_of_available_dancers = len(available_dancers)
    if number_of_available_dancers > 0:
        random_order = random.sample(range(0, number_of_available_dancers), number_of_available_dancers)
        for num in range(len(random_order)):
            dancer = available_dancers[random_order[num]]
            dancer_id = dancer[gen_dict[sql_id]]
            if dancer_id is not None:
                query = ' SELECT * FROM {tn} WHERE {id} = ?'.format(tn=selection_list, id=sql_id)
                dancer_available = curs.execute(query, (dancer_id,)).fetchone()
                if dancer_available is not None:
                    partner_id = find_partner(dancer_id, connection=conn, cursor=curs)
                    if partner_id is not None:
                        create_pair(dancer_id, partner_id, connection=conn, cursor=curs)
                        move_selected_contestant(dancer_id, connection=conn, cursor=curs)
                        move_selected_contestant(partner_id, connection=conn, cursor=curs)
            query = 'SELECT * FROM {tn}'.format(tn=selected_list)
            remaining_dancers = curs.execute(query).fetchall()
            number_of_selected_dancers = len(remaining_dancers) + buffer_for_selection
            if number_of_selected_dancers >= max_contestants:
                break
    conn.commit()
    reset_selection_tables(connection=conn, cursor=curs)
    status_print('')
    status_print("--- Done in %.3f seconds ---" % (time.time() - start_time))
    status_print('')

    ####################################################################################################################
    # Create signup excel file
    ####################################################################################################################
    # otp = selection.replace(".", ("_"+str(start_time)+"."))
    # wb = openpyxl.Workbook()
    # wb.save(otp)
    # if exists(selection):
    #     wb = openpyxl.load_workbook(selection)
    # query = 'SELECT * FROM {tn} ORDER BY {id}'.format(tn=selected_list, id=sql_id)
    # selected_contestants = curs.execute(query).fetchall()

    ####################################################################################################################
    # Collect user data from main selection
    ####################################################################################################################
    collect_city_overview(source_table=fixed_beginners_list, target_table=beginners_list, users=Beginners,
                          cursor=curs, connection=conn)
    collect_city_overview(source_table=fixed_lions_list, target_table=lions_list, users=Lions,
                          cursor=curs, connection=conn)
    collect_city_overview(source_table=selected_list, target_table=contestants_list, users=contestants,
                          cursor=curs, connection=conn)

    ####################################################################################################################
    # Collect individual data for statistical analysis
    ####################################################################################################################
    if debug:
        query = 'SELECT * FROM {tn}'.format(tn=signup_list)
        all_dancers = curs.execute(query).fetchall()
        query = 'SELECT name FROM sqlite_master WHERE type = "table" AND name = "{tn}"'.format(tn=individual_list)
        individual_table_exists = len(curs.execute(query).fetchall())
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
        last_run = len(curs.execute(query).fetchall())
        this_run = last_run + 1
        query = 'INSERT INTO {tn} ({run}) VALUES (?)'.format(tn=individual_list, run=sql_run)
        curs.execute(query, (this_run,))
        for dancer in all_dancers:
            dancer_id = dancer[gen_dict[sql_id]]
            query = 'SELECT * FROM {tn} WHERE {id} = ?'.format(id=sql_id, tn=selected_list)
            dancer_selected = curs.execute(query, (dancer_id,)).fetchall()
            if len(dancer_selected) == 0:
                dancer_selected = 0
            else:
                dancer_selected = 1
            query = 'UPDATE {tn} SET "{col}" = ? WHERE {run} = ?'.format(tn=individual_list, run=sql_run, col=dancer_id)
            curs.execute(query, (dancer_selected, this_run))
        conn.commit()

    # Update status
    status_update(cursor=curs)

    # Close cursor and connection
    curs.close()
    conn.close()

if __name__ == "__main__":
    root = Tk()
    root.geometry("1820x900")
    root.state('zoomed')
    # root.option_add("*Font", tkFont.nametofont('TkFixedFont'))
    pad_out = 8
    pad_in = 8
    frame = Frame()
    frame.place(in_=root, anchor="c", relx=.50, rely=.50)
    xscrollbar = Scrollbar(master=frame, orient=HORIZONTAL)
    xscrollbar.grid(row=2, column=0, padx=pad_in, sticky=E+W)
    status_text = Text(master=frame, width=148, height=50, padx=pad_in, pady=pad_in,
                       xscrollcommand=True, yscrollcommand=True, state=DISABLED, wrap=NONE)
    status_text.grid(row=0, column=0, padx=pad_out, rowspan=2)
    xscrollbar.config(command=status_text.xview)
    cli_text = Entry(master=frame, width=200)
    cli_text.grid(row=3, column=0, padx=pad_in, pady=pad_out)
    cli_text.bind('<Return> ', cli_parser)
    data_help_frame = Frame(master=frame)
    data_help_frame.grid(row=0, column=1, rowspan=4, columnspan=3)
    data_text = Text(master=data_help_frame, width=70, height=30, padx=pad_in, pady=pad_in, wrap=WORD, state=DISABLED)
    data_text.grid(row=0, column=0, padx=pad_out, columnspan=3)
    padding_frame = Frame(master=data_help_frame, height=16)
    padding_frame.grid(row=1, column=0)
    help_text = Text(master=data_help_frame, width=70, height=19, padx=pad_in, pady=pad_in, wrap=WORD, state=DISABLED)
    help_text.grid(row=2, column=0, padx=pad_out, columnspan=3)
    start_button = Button(master=data_help_frame, text='Start new selection database', command=main_selection)
    start_button.grid(row=3, column=0, padx=pad_out, pady=pad_in)
    update_button = Button(master=data_help_frame, text='Select existing database')
    update_button.grid(row=3, column=1, padx=pad_out)
    options_button = Button(master=data_help_frame, text='Options')
    options_button.grid(row=3, column=2, padx=pad_out)
    welcome_text()
    print_ntds_config()
    cli_text.focus_set()
    if selected_database != '':
        main_conn = sql.connect(selected_database)
        main_curs = main_conn.cursor()
        status_update(cursor=main_curs)
        main_curs.close()
        main_conn.close()
    root.mainloop()
