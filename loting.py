# Main file for the program
import sqlite3 as sql
import openpyxl
# from openpyxl import Workbook
from random import randint
import random
# from collections import Counter
import time
import logging
logging.basicConfig(filename='loting.log', filemode='w', level=logging.DEBUG)

testing = True

# init stuff
participating_teams = 'deelnemende_teams.xlsx'
template = 'NTDS_Template.xlsx'

# boundaries
max_contestants = 400
max_fixed_beginner_pairs = 2
max_fixed_beginners = max_fixed_beginner_pairs*2
max_fixed_lion_pairs = 5
max_fixed_lion_contestants = max_fixed_lion_pairs*2
buffer_for_selection = 40

# Names
database_name = 'main_data.db'
signup_list = 'signup_list'
selected_list = 'selected_people'
team_list = 'team_list'
partners_selection_pool = 'partners_selection_pool'
partners_list = 'partners_list'
preselected_people_list = 'preselected_people'

# more names
breiten = 'Breiten'
beginner = 'Beginner'
open_class = 'Open'

# SQL Table column names and dictionary for dancers lists
sql_id = 'id'
sql_name = 'name'
sql_email = 'email'
sql_bn = 'ballroom_level'
sql_ln = 'latin_class'
sql_bp = 'ballroom_partner'
sql_lp = 'latin_partner'
sql_rol = 'role'
sql_br = 'ballroom_role'
sql_lr = 'latin_role'
sql_bbd = 'ballroom_mandatory_blind_date'
sql_lbd = 'latin_mandatory_blind_date'
sql_tc = 'team_captain'
sql_team = 'team'
gen_dict = {sql_id: 0, sql_name: 1, sql_email: 2, sql_bn: 3, sql_ln: 4, sql_bp: 5, sql_lp: 6, sql_br: 7, sql_lr: 8,
            sql_bbd: 9, sql_lbd: 10, sql_tc: 11, sql_team: 12}

# SQL Table column names and dictionary of teams list
sql_city = 'city'
sql_sl = 'signup_list'
team_dict = {sql_team: 0, sql_city: 1, sql_sl: 2}

# SQL Table column names and dictionary for partners list
sql_lead = 'lead'
sql_follow = 'follow'
sql_city_lead = 'city_lead'
sql_city_follow = 'city_follow'
p_dict = {sql_lead: 0, sql_follow: 1, sql_city_lead: 2, sql_city_follow: 3}

# General query formats
drop_table_query = 'DROP TABLE IF EXISTS {};'
dancers_list_query = 'CREATE TABLE {tn} ({id} INT PRIMARY KEY, {name} TEXT, {email} TEXT, {bn} TEXT, {ln} TEXT,' \
                     ' {bp} INT, {lp} INT, {br} TEXT, {lr} TEXT, {bbd} TEXT, {lbd} TEXT, {tc} TEXT, {team} TEXT);'\
                     .format(tn={}, id=sql_id, name=sql_name, email=sql_email, bn=sql_bn, ln=sql_ln, bp=sql_bp,
                             lp=sql_lp, br=sql_br, lr=sql_lr, bbd=sql_bbd, lbd=sql_lbd, tc=sql_tc, team=sql_team)
team_list_query = 'CREATE TABLE {tn} ({team} TEXT PRIMARY KEY, {city} TEXT, {signup_list} TEXT);' \
    .format(tn={}, team=sql_team, city=sql_city, signup_list=sql_sl)
paren_table_query = 'CREATE TABLE {tn} ({lead} INT PRIMARY KEY, {follow} INT, {lead_city} TEXT, {follow_city} TEXT);' \
    .format(tn={}, lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow)


def get_team(identifier, connection, cursor):
    """Finds the team of a given id"""
    query = 'SELECT {team} FROM {tn} WHERE {identifier} =?'.format(tn=signup_list, identifier=sql_id, team=sql_team)
    team = cursor.execute(query, (identifier,)).fetchone()[gen_dict[sql_id]]
    connection.commit()
    return team


def move_selected_contestant(identifier, connection, cursor):
    """Moves dancer, given id, from the signup list to selected list"""
    if identifier is not None:
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE id = ?;'.format(tn1=selected_list, tn2=signup_list)
        cursor.execute(query, (identifier,))
        query = 'DELETE FROM {tn} WHERE id = ?'.format(tn=signup_list)
        cursor.execute(query, (identifier,))
        connection.commit()
        logging.info('Selected {} for the NTDS'.format(identifier))


def copy_to_selection_pool(selected_people, connection, cursor):
    """Copies a number of dancers to the selection pool, given a list of id's"""
    list_of_ids = [(row[gen_dict[sql_id]],) for row in selected_people]
    query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
        .format(tn1=partners_selection_pool, tn2=signup_list, id=sql_id)
    cursor.executemany(query, list_of_ids)
    connection.commit()


def copy_to_preselection(preselected_people, connection, cursor):
    """Copies a number of dancers to the selection pool, given a list of people"""
    list_of_ids = [(row[gen_dict[sql_id]],) for row in preselected_people]
    query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
        .format(tn1=preselected_people_list, tn2=signup_list, id=sql_id)
    cursor.executemany(query, list_of_ids)
    # query = 'DELETE FROM {tn1} WHERE {id} = ?;'.format(tn1=signup_list, id=sql_id)
    # cursor.executemany(query, list_of_ids)
    connection.commit()


def find_signed_partner(dancer, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    dancer_id = dancer[gen_dict[sql_id]]
    partner_id = None
    ballroom_partner = dancer[gen_dict[sql_bp]]
    latin_partner = dancer[gen_dict[sql_lp]]
    team = dancer[gen_dict[sql_team]]
    # Check if the contestant already has signed up with a partner
    if isinstance(ballroom_partner, int):
        partner_id = ballroom_partner
    if all([isinstance(latin_partner, int), partner_id is None]):
        partner_id = latin_partner
    # query = 'SELECT * from {tn} WHERE {team} != ?'.format(tn=preselected_people_list, team=sql_team)
    # potential_partners = cursor.execute(query).fetchall()
    # if len(potential_partners) > 0:
    #     print('')
    if partner_id is None:
        logging.info('{id1} has not signed with a partner'.format(id1=dancer_id))
    else:
        logging.info('Matched {id1} and {id2} together'.format(id1=dancer_id, id2=partner_id))
    return partner_id


def find_partner(identifier, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=signup_list, id=sql_id)
    data = cursor.execute(query, (identifier,)).fetchone()
    ballroom_level = data[gen_dict[sql_bn]]
    latin_level = data[gen_dict[sql_ln]]
    ballroom_partner = data[gen_dict[sql_bp]]
    latin_partner = data[gen_dict[sql_lp]]
    rol = data[gen_dict[sql_br]]
    # ballroom_mandatory_date = data[gen_dict[sql_bbd]]
    # latin_mandatory_date = data[gen_dict[sql_lbd]]
    # team_captain = data[gen_dict[sql_tc]]
    team = data[gen_dict[sql_team]]
    # Check if the contestant already has signed up with a partner
    if isinstance(ballroom_partner, int):
        partner_id = ballroom_partner
    if all([isinstance(latin_partner, int), partner_id is None]):
        partner_id = latin_partner
    # If the contestant is a beginner, and has no partner, find a partner
    # schrijf speciale select beginner
    if all([ballroom_level == beginner, latin_level == beginner, partner_id is None]):
        query = 'SELECT * FROM {tn} where {bn} = ? AND {ln} = ? AND {bp} = "" AND {lp} = "" AND {br} != ? ' \
                'AND {team} != ?' \
            .format(tn=partners_selection_pool, bn=sql_bn, ln=sql_ln, bp=sql_bp, lp=sql_lp, br=sql_br, team=sql_team)
        potential_partners = cursor.execute(query, (ballroom_level, latin_level, rol, team)).fetchall()
        number_of_potential_partners = len(potential_partners)
        if number_of_potential_partners > 0:
            random_num = randint(0, number_of_potential_partners - 1)
            partner_id = potential_partners[random_num][gen_dict[sql_id]]
    # Try to find the best partner for someone that has either Breiten or Open level

    # if all([ballroom_level == latin_level, partner_id is None]):
    #     query = 'SELECT * FROM {tn} where {bn} = ? AND {ln} = ? AND {bp} = "" AND {lp} = "" AND {br} != ? ' \
    #             'AND {team} != ?'\
    #         .format(tn=signupList, bn=sql_bn, ln=sql_ln, bp=sql_bp, lp=sql_lp, br=sql_br, team=sql_team)
    #     potential_partners = curs.execute(query,(ballroom_level, latin_level, rol, team)).fetchall()
    #     number_of_potential_partners = len(potential_partners)
    #     random_id = randint(0, number_of_potential_partners - 1)
    #     partner_id = potential_partners[random_id][0]
    #     return partner_id
    # if ballroom_level in [Beginner, Breiten]:
    #     niveau = ballroom_level
    #
    #     query = 'SELECT * FROM {tn} where {bn} = ? AND {bp} = "" AND {lp} = "" AND {br} != ? AND {team} != ?' \
    #         .format(tn=signupList, bn=sql_bn, bp=sql_bp, lp=sql_lp, br=sql_br, team=sql_team)
    #     potential_partners = curs.execute(query, (ballroom_level, rol, team)).fetchall()
    #     number_of_potential_partners = len(potential_partners)
    #     random_id = randint(0, number_of_potential_partners - 1)
    #     partner_id = potential_partners[random_id][0]
    #     return partner_id
    connection.commit()
    if partner_id is None:
        logging.info('Found no match for {id1}'.format(id1=identifier))
    else:
        logging.info('Matched {id1} and {id2} together'.format(id1=identifier, id2=partner_id))
    return partner_id


def find_beginner_partner(identifier, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=signup_list, id=sql_id)
    data = cursor.execute(query, (identifier,)).fetchone()
    ballroom_level = data[gen_dict[sql_bn]]
    latin_level = data[gen_dict[sql_ln]]
    ballroom_partner = data[gen_dict[sql_bp]]
    latin_partner = data[gen_dict[sql_lp]]
    rol = data[gen_dict[sql_br]]
    # ballroom_mandatory_date = data[gen_dict[sql_bbd]]
    # latin_mandatory_date = data[gen_dict[sql_lbd]]
    # team_captain = data[gen_dict[sql_tc]]
    team = data[gen_dict[sql_team]]
    # Check if the contestant already has signed up with a partner
    if isinstance(ballroom_partner, int):
        partner_id = ballroom_partner
    if all([isinstance(latin_partner, int), partner_id is None]):
        partner_id = latin_partner
    # If the contestant is a beginner, and has no partner, find a partner
    # schrijf speciale select beginner
    if all([ballroom_level == beginner, latin_level == beginner, partner_id is None]):
        query = 'SELECT * FROM {tn} where {bn} = ? AND {ln} = ? AND {bp} = "" AND {lp} = "" AND {br} != ? ' \
                'AND {team} != ?' \
            .format(tn=signup_list, bn=sql_bn, ln=sql_ln, bp=sql_bp, lp=sql_lp, br=sql_br, team=sql_team)
        potential_partners = cursor.execute(query, (ballroom_level, latin_level, rol, team)).fetchall()
        number_of_potential_partners = len(potential_partners)
        if number_of_potential_partners > 0:
            random_num = randint(0, number_of_potential_partners - 1)
            partner_id = potential_partners[random_num][gen_dict[sql_id]]
    connection.commit()
    if partner_id is None:
        logging.info('Found no match for {id1}'.format(id1=identifier))
    else:
        logging.info('Matched {id1} and {id2} together'.format(id1=identifier, id2=partner_id))
    return partner_id


def create_pair(lead, follow, connection, cursor):
    """Temp"""
    lead_team = get_team(lead, connection=connection, cursor=cursor)
    follow_team = get_team(follow, connection=connection, cursor=cursor)
    query = 'INSERT INTO {tn} VALUES (?, ?, ?, ?)'.format(tn=partners_list)
    cursor.execute(query, (lead, follow, lead_team, follow_team))
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=partners_selection_pool, id=sql_id)
    cursor.executemany(query, [(lead,), (follow,)])
    connection.commit()


def no_partner_found(identifier, connection, cursor):
    """Temp"""
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=partners_selection_pool, id=sql_id)
    cursor.execute(query, (identifier,))
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
    tables_to_drop = [signup_list, selected_list, team_list,
                      partners_list, partners_selection_pool, preselected_people_list]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        cursor.execute(query)
    # Create new tables
    dancer_list_tables = [signup_list, selected_list, partners_selection_pool, preselected_people_list]
    for item in dancer_list_tables:
        query = dancers_list_query.format(item)
        cursor.execute(query)
    query = team_list_query.format(team_list)
    cursor.execute(query)
    query = paren_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()


def create_competing_teams(connection, cursor):
    """"Creates list of all competing cities"""
    # Create Workbook object and a Worksheet from it
    wb = openpyxl.load_workbook(participating_teams)
    ws = wb.worksheets[0]
    # Empty 2d list that will contain the signup sheet filenames for each of the teams
    competing_cities_array = []
    # Fill up list with competing cities
    for i in range(2, ws.max_row + 1):
        competing_cities_array \
            .append([ws.cell(row=i, column=1).value, ws.cell(row=i, column=2).value, ws.cell(row=i, column=3).value])
    query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=team_list)
    for row in competing_cities_array:
        cursor.execute(query, row)
    connection.commit()
    return competing_cities_array


def main():
    start_time = time.time()
    # Connect to database and create a cursor
    conn = sql.connect(database_name)
    curs = conn.cursor()
    # Create SQL tables
    create_tables(connection=conn, cursor=curs)
    # Create competing cities list
    competing_teams = create_competing_teams(connection=conn, cursor=curs)
    competing_cities = [row[team_dict[sql_city]] for row in competing_teams]
    number_of_competing_cities = len(competing_cities)
    # Get maximum number of columns
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    max_col = max_rc('col', ws)
    # Copy the signup list from every team into the SQL database
    total_signup_list = list()
    total_number_of_contestants = 0
    for team in competing_teams:
        city = team[team_dict[sql_city]]
        team_signup_list = team[team_dict[sql_sl]]
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
        query = 'INSERT INTO {} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?);'.format(signup_list)
        for row in city_signup_list:
            curs.execute(query, row)
        total_number_of_contestants += max_row-1
    conn.commit()

    # Select the team captains
    query = 'SELECT * FROM {tn1} WHERE {tc} = "Ja"'.format(tn1=signup_list, tc=sql_tc)
    team_captains = curs.execute(query).fetchall()
    for captain in team_captains:
        captain_id = captain[gen_dict[sql_id]]
        partner_id = find_partner(captain_id, connection=conn, cursor=curs)
        move_selected_contestant(captain_id, connection=conn, cursor=curs)
        move_selected_contestant(partner_id, connection=conn, cursor=curs)
    conn.commit()

    # Select the beginners of cities that have less than the maximum guaranteed
    guaranteed_beginners = []
    for city in competing_cities:
        query = 'SELECT * FROM {tn1} WHERE {bn} = ? AND {ln} = ? AND {team} = ?'\
            .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team)
        city_beginners = curs.execute(query, (beginner, beginner, city)).fetchall()
        number_of_city_beginners = len(city_beginners)
        if number_of_city_beginners <= max_fixed_beginners:
            guaranteed_beginners.extend(city_beginners)
    copy_to_preselection(guaranteed_beginners, connection=conn, cursor=curs)
    # If the beginners have signed with a partner, group them and place them in the preselected pairs
    # If the beginners have not signed with a partner, place them in the preselected beginners
    for beg in guaranteed_beginners:
        beginner_id = beg[gen_dict[sql_id]]
        partner_id = find_signed_partner(beg, connection=conn, cursor=curs)
        if partner_id is not None:
            create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
    # Transfer all beginners to selection pool and get a list of beginners from the city
    query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}"' \
        .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner)
    all_beginners = curs.execute(query).fetchall()
    copy_to_selection_pool(all_beginners, connection=conn, cursor=curs)
    for beg in guaranteed_beginners:
        beginner_id = beg[gen_dict[sql_id]]
        partner_id = find_partner(beg, connection=conn, cursor=curs)
        if partner_id is not None:
            create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
    query = 'SELECT * FROM {tn}'.format(tn=partners_selection_pool)
    all_beginners = curs.execute(query).fetchall()
    number_of_beginners = len(all_beginners)
    while number_of_beginners > 0:
        random_int = randint(0, number_of_beginners - 1)
        beginner_id = all_beginners[random_int][gen_dict[sql_id]]
        partner_id = find_partner(beginner_id, connection=conn, cursor=curs)
        if partner_id is None:
            partner_id = find_beginner_partner(beginner_id, connection=conn, cursor=curs)
        partner_team = get_team(partner_id, connection=conn, cursor=curs)
        if partner_id is not None:
            if all_beginners[random_int][gen_dict[sql_br]] == 'Lead':
                create_pair(lead=beginner_id, follow=partner_id, connection=conn, cursor=curs)
            else:
                create_pair(lead=partner_id, follow=beginner_id, connection=conn, cursor=curs)
        else:
            no_partner_found(beginner_id, connection=conn, cursor=curs)
        query = 'SELECT * FROM {tn}'.format(tn=partners_selection_pool)
        all_beginners = curs.execute(query).fetchall()
        number_of_beginners = len(all_beginners)
    # Get beginners
    max_number_of_beginners = 0
    for city in competing_cities:
        query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}" AND {team} = "{city}"' \
            .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team, city=city)
        if len(curs.execute(query).fetchall()) >= max_fixed_beginners:
            max_number_of_beginners += max_fixed_beginners
        else:
            max_number_of_beginners += len(curs.execute(query).fetchall())
    max_number_of_beginners += max_number_of_beginners % 2
    max_number_of_pairs = int(max_number_of_beginners / 2)

    required_tables = [partners_list, partners_selection_pool]
    for item in required_tables:
        query = drop_table_query.format(item)
        curs.execute(query)
    query = dancers_list_query.format(partners_selection_pool)
    curs.execute(query)
    query = paren_table_query.format(partners_list)
    curs.execute(query)
    conn.commit()
    # Transfer all beginners to selection pool and get a list of beginners from the city
    query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}"' \
        .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner)
    all_beginners = curs.execute(query).fetchall()
    copy_to_selection_pool(all_beginners, connection=conn, cursor=curs)
    number_of_beginners = len(all_beginners)
    while number_of_beginners > 0:
        random_int = randint(0, number_of_beginners - 1)
        beginner_id = all_beginners[random_int][gen_dict[sql_id]]
        partner_id = find_partner(beginner_id, connection=conn, cursor=curs)
        if partner_id is not None:
            if all_beginners[random_int][gen_dict[sql_br]] == 'Lead':
                create_pair(lead=beginner_id, follow=partner_id, connection=conn, cursor=curs)
            else:
                create_pair(lead=partner_id, follow=beginner_id, connection=conn, cursor=curs)
        else:
            no_partner_found(beginner_id, connection=conn, cursor=curs)
        query = 'SELECT * FROM {tn}'.format(tn=partners_selection_pool)
        all_beginners = curs.execute(query).fetchall()
        number_of_beginners = len(all_beginners)

    ###################
    # NEW CODE 2
    ###################
    # query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}"'\
    #     .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner)
    # all_beginners = curs.execute(query).fetchall()
    # number_of_beginners = len(all_beginners)
    # selection_cities = []

    selected_pairs = []
    beginners_selected = False
    while beginners_selected is False:
        query = 'SELECT * FROM {tn} ORDER BY {lead}'.format(tn=partners_list, lead=sql_lead, follow=sql_follow)
        all_pairs = curs.execute(query).fetchall()
        number_of_pairs = len(all_pairs)
        dummy = 0
        selected_cities = []
        selected_pairs = []
        random_numbers = random.sample(range(0, number_of_pairs - 1), max_number_of_pairs)
        for num in random_numbers:
            selected_cities.append(all_pairs[num][2])
            selected_cities.append(all_pairs[num][3])
            selected_pairs.append(all_pairs[num][0])
            selected_pairs.append(all_pairs[num][1])
        for city in competing_cities:
            query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}" AND {team} = "{city}"' \
                .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team, city=city)
            max_number_of_city_beginners = len(curs.execute(query).fetchall())
            if max_number_of_city_beginners > max_fixed_beginners:
                max_number_of_city_beginners = max_fixed_beginners
            number_of_city_beginners = selected_cities.count(city)
            if number_of_city_beginners == max_number_of_city_beginners or number_of_city_beginners >= max_fixed_beginners:
                dummy += 1
        if dummy == number_of_competing_cities:
            beginners_selected = True
    for contestant in selected_pairs:
        move_selected_contestant(contestant, connection=conn, cursor=curs)
    print('dummy')
    # for pair in selected_pairs:
    #     lead_id = pair[p_dict[sql_lead]]
    #     follow_id = pair[p_dict[sql_follow]]
    #     move_selected_contestant(lead_id, connection=conn, cursor=curs)
    #     move_selected_contestant(follow_id, connection=conn, cursor=curs)

    ###################
    # NEW CODE 3
    ###################
    # for city in competing_cities:
    #     query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}" AND {team} = "{city}"' \
    #         .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team, city=city)
    #     city_beginners = curs.execute(query).fetchall()
    #     number_of_city_beginners = len(city_beginners)
    #     if number_of_city_beginners <= max_fixed_beginners:
    #         for contestant in city_beginners:
    #             contestant_id = contestant[gen_dict[sql_id]]
    #             partner_id = find_partner(contestant_id, connection=conn, cursor=curs)
    #             move_selected_contestant(contestant_id, connection=conn, cursor=curs)
    #             move_selected_contestant(partner_id, connection=conn, cursor=curs)
    #     else:
    #         print('dummy')

    # ###################
    # # NEW CODE
    # ###################
    # query = 'SELECT * FROM {tn} ORDER BY {lead}'.format(tn=partners_list, lead=sql_lead, follow=sql_follow)
    # all_pairs = curs.execute(query).fetchall()
    # number_of_beginners = len(all_pairs)
    # completed_cities_list = []
    # beginners_selected = False
    # while beginners_selected is False:
    #     for row in competing_cities_list:
    #         city = row[team_dict[sql_city]]
    #         if city not in completed_cities_list:
    #             filt = ''
    #             for team in completed_cities_list:
    #                 filt += ' AND ({city_lead} NOT LIKE "{team}" AND {city_follow} NOT LIKE "{team}")'\
    #                     .format(city_lead=sql_city_lead, city_follow=sql_city_follow, team=team)
    #             query = 'SELECT * FROM {tn1} WHERE {city_lead} = "{city}" OR {city_follow} = "{city}"'\
    #                 .format(tn1=partners_list, city_lead=sql_city_lead, city_follow=sql_city_follow, city=city)
    #             query += filt
    #             pairs = curs.execute(query).fetchall()
    #             if len(pairs) == 0:
    #                 query = 'SELECT * FROM {tn1} WHERE {city_lead} = "{city}" OR {city_follow} = "{city}"' \
    #                     .format(tn1=partners_list, city_lead=sql_city_lead, city_follow=sql_city_follow, city=city)
    #                 pairs = curs.execute(query).fetchall()
    #             number_of_beginners = len(pairs)
    #             random_number = randint(0, number_of_beginners-1)
    #             selected_pair = pairs[random_number]
    #             lead_id = selected_pair[p_dict[sql_lead]]
    #             follow_id = selected_pair[p_dict[sql_follow]]
    #             move_selected_contestant(lead_id, connection=conn, cursor=curs)
    #             move_selected_contestant(follow_id, connection=conn, cursor=curs)
    #             query = 'DELETE FROM {tn} WHERE {lead} = ?'.format(tn=partners_list, lead=sql_lead)
    #             curs.execute(query, (lead_id,))
    #             conn.commit()
    #             query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"'\
    #                 .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=city)
    #             number_of_selected_beginners = len(curs.execute(query).fetchall())
    #             if number_of_selected_beginners >= max_fixed_beginners:
    #                 completed_cities_list.append(city)
    #         if len(completed_cities_list) == len(competing_cities_list):
    #             beginners_selected = True


    # for i in range (0, len(competing_cities_list)):
    #     # SELECT * FROM partners_list WHERE (city_lead not like "Groningen" AND city_follow not like "Groningen") AND (city_lead not like "Utrecht" AND city_follow not like "Utrecht")
    #     print('dummy')
    #     beginners_not_selected_yet = False
        # for i in range(0, number_of_competing_cities):
        #     # Temp
        #     query = 'SELECT min({id}) FROM {tn1} WHERE {team} = "{city}"' \
        #         .format(id=sql_id, tn1=partners_selection_pool, team=sql_team, city=city)
        #     min_id = curs.execute(query).fetchone()[gen_dict[sql_id]]
        #     # Get beginners of city
        #     query = 'SELECT * FROM {tn1} WHERE {team} = "{city}"' \
        #         .format(tn1=partners_selection_pool, team=sql_team, city=city)
        #     beginners_city = curs.execute(query).fetchall()
        #     number_of_beginners = len(beginners_city)
        #     # min_id = 0
        #     while number_of_beginners > 0:
        #         random_int = randint(0, number_of_beginners - 1)
        #         beginner_id = beginners_city[random_int][gen_dict[sql_id]]
        #         partner_id = find_partner(beginner_id, connection=conn, cursor=curs)
        #         if partner_id is not None:
        #             if beginners_city[random_int][gen_dict[sql_br]] == 'Lead':
        #                 create_paar(lead=beginner_id, follow=partner_id, connection=conn, cursor=curs)
        #             else:
        #                 create_paar(lead=partner_id, follow=beginner_id, connection=conn, cursor=curs)
        #             query = 'SELECT * FROM {tn} WHERE {team} = "{city}"'\
        #                 .format(tn=partners_selection_pool, team=sql_team, city=city)
        #         else:
        #             no_partner_found(beginner_id, connection=conn, cursor=curs)
        #         beginners_city = curs.execute(query).fetchall()
        #         number_of_beginners = len(beginners_city)
        #     query = 'SELECT * FROM {tn} WHERE {lead} >= ? AND {follow} >= ? ORDER BY {lead}'\
        #         .format(tn=partners_list, lead=sql_lead, follow=sql_follow)
        #     paren = curs.execute(query, [(min_id), (min_id)]).fetchall()
        #     aantal_paren = len(paren)
        #     query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"' \
        #         .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=city)
        #     number_of_selected_beginners = len(curs.execute(query).fetchall())
        #     # if aantal_paren + number_of_selected_beginners < max_fixed_beginners:
        #     #     print('dummy')
        #     #     teams_paren = []
        #     #     for row in paren:
        #     #         teams_paren.append(get_team(row[0], connection=conn, cursor=curs))
        #     #         teams_paren.append(get_team(row[1], connection=conn, cursor=curs))
        #     #     for contestant in teams_paren:
        #     #         move_selected_contestant(contestant, connection=conn, cursor=curs)
        #     #     query = 'SELECT * FROM {tn} WHERE {lead} >= ? AND {follow} >= ? ORDER BY {lead}' \
        #     #         .format(tn=partners_list, lead=sql_lead, follow=sql_follow)
        #     #     paren = curs.execute(query, [(min_id), (min_id)]).fetchall()
        #     #     aantal_paren = len(paren)
        #     #     query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"' \
        #     #         .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=city)
        #     #     number_of_selected_beginners = len(curs.execute(query).fetchall())
        #     if aantal_paren <= max_fixed_beginners:
        #         random_numbers = list(range(aantal_paren))
        #     if aantal_paren > max_fixed_beginners:
        #         random_numbers = random.sample(range(0, aantal_paren-1), max_fixed_beginners)
        #     geselecteerde_paren = [paren[random_numbers[i]] for i in range(0, len(random_numbers))]
        #     teams_paren = []
        #     for j in range(len(random_numbers)):
        #         teams_paren.append(get_team(geselecteerde_paren[j][0], connection=conn, cursor=curs))
        #         teams_paren.append(get_team(geselecteerde_paren[j][1], connection=conn, cursor=curs))
        #         if teams_paren.count(city) + number_of_selected_beginners >= max_fixed_beginners:
        #             break
        #     geselecteerde_paren = geselecteerde_paren[0:int(len(teams_paren) / 2)]
        #     geselecteerde_paren = [j for k in geselecteerde_paren for j in k]
        #     for j in range(len(geselecteerde_paren)):
        #         move_selected_contestant(geselecteerde_paren[j], connection=conn, cursor=curs)

    # wb = Workbook()
    # ws = wb.active
    # for i in range (0, len(total_signup_list)):
    #     for j in range (0, len(total_signup_list[0])):
    #         ws.cell(row = i+1,column = j+1).value = total_signup_list[i][j].value
    #         # ws.column_dimensions[get_column_letter(i + 1)].width = 30
    # wb.save(filename='example.xlsx')

    # selectie
    # teamcaptain (+partner) selecteren
    # twee beginnerparen selecteren
    # vijf paren voor leeuw selecteren
    # rest gaat in lotingspoel

    if testing:
        query = 'CREATE TABLE IF NOT EXISTS "beginners" (id INTEGER PRIMARY KEY AUTOINCREMENT, {c1} INT, {c2} INT, {c3} INT, {c4} INT)'\
            .format(c1='Enschede', c2='Eindhoven', c3='Groningen', c4='Utrecht')
        results = []
        curs.execute(query)
        temp = []
        for i in range(len(competing_teams)):
            team = competing_teams[i][team_dict[sql_city]]
            query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"' \
                .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=team)
            selected_beginners = curs.execute(query).fetchall()
            number_selected = len(selected_beginners)
            overview = 'Number of selected beginners from {city} is: {number}'.format(city=team, number=number_selected)
            print(overview)
            temp.append(number_selected)
        results.append(temp)
        query = 'INSERT INTO "beginners" ({c1},{c2},{c3},{c4}) VALUES (?,?,?,?)'.format(c1='Enschede', c2='Eindhoven', c3='Groningen', c4='Utrecht')
        curs.execute(query, (results[0][0], results[0][1], results[0][2], results[0][3]))
        conn.commit()

    # Close cursor and connection
    curs.close()
    conn.close()
    print("--- %s seconds ---" % (time.time() - start_time))
    print('Done!')

if __name__ == '__main__':
    for i in range(10):
        main()
