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
# logger = logging.getLogger()
# logger.setLevel(logging.DEBUG)
# logger.addHandler(logging.StreamHandler())

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
selection_list = 'selection_list'
selected_list = 'selected_people'
team_list = 'team_list'
partners_selection_pool = 'partners_selection_pool'
partners_list = 'partners_list'
ref_partner_list = 'reference_partner_list'
preselected_people_list = 'preselected_people'
fixed_beginners_list = 'fixed_beginners'
fixed_lions_list = 'fixed_lions'

# more names
breiten = 'Breiten'
beginner = 'Beginner'
open_class = 'Open'

# Participants Lion points
lion_participants = [beginner, breiten]

# SQL Table column names and dictionary for dancers lists
sql_id = 'id'
sql_name = 'name'
sql_email = 'email'
sql_bl = 'ballroom_class'
sql_ll = 'latin_class'
sql_bp = 'ballroom_partner'
sql_lp = 'latin_partner'
sql_rol = 'role'
sql_br = 'ballroom_role'
sql_lr = 'latin_role'
sql_bbd = 'ballroom_mandatory_blind_date'
sql_lbd = 'latin_mandatory_blind_date'
sql_tc = 'team_captain'
sql_city = 'city'
gen_dict = {sql_id: 0, sql_name: 1, sql_email: 2, sql_bl: 3, sql_ll: 4, sql_bp: 5, sql_lp: 6, sql_br: 7, sql_lr: 8,
            sql_bbd: 9, sql_lbd: 10, sql_tc: 11, sql_city: 12}

# SQL Table column names and dictionary of teams list
sql_team = 'team'
sql_sl = 'signup_list'
team_dict = {sql_team: 0, sql_city: 1, sql_sl: 2}

# SQL Table column names and dictionary for partners list
sql_lead = 'lead'
sql_follow = 'follow'
sql_city_lead = 'city_lead'
sql_city_follow = 'city_follow'
partner_dict = {sql_lead: 0, sql_follow: 1, sql_city_lead: 2, sql_city_follow: 3}

# SQL Table column names and dictionary for city list
sql_num_con = 'number_of_contestants'
sql_max_con = 'max_contestants'
city_dict = {sql_city: 0, sql_num_con: 1, sql_max_con: 2}

# General query formats
drop_table_query = 'DROP TABLE IF EXISTS {};'
dancers_list_query = 'CREATE TABLE {tn} ({id} INT PRIMARY KEY, {name} TEXT, {email} TEXT, {bl} TEXT, {ln} TEXT,' \
                     ' {bp} INT, {lp} INT, {br} TEXT, {lr} TEXT, {bbd} TEXT, {lbd} TEXT, {tc} TEXT, {city} TEXT);'\
                     .format(tn={}, id=sql_id, name=sql_name, email=sql_email, bl=sql_bl, ln=sql_ll, bp=sql_bp,
                             lp=sql_lp, br=sql_br, lr=sql_lr, bbd=sql_bbd, lbd=sql_lbd, tc=sql_tc, city=sql_city)
team_list_query = 'CREATE TABLE {tn} ({team} TEXT PRIMARY KEY, {city} TEXT, {signup_list} TEXT);' \
    .format(tn={}, team=sql_team, city=sql_city, signup_list=sql_sl)
paren_table_query = 'CREATE TABLE {tn} ({lead} INT, {follow} INT, {lead_city} TEXT, {follow_city} TEXT);' \
    .format(tn={}, lead=sql_lead, follow=sql_follow, lead_city=sql_city_lead, follow_city=sql_city_follow)
city_list_query = 'CREATE TABLE {tn} ({city} TEXT, {beg} INT, {max_beg} INT);' \
    .format(tn={}, city=sql_city, beg=sql_num_con, max_beg=sql_max_con)


def get_team(identifier, cursor):
    """Finds the team of a given id"""
    query = 'SELECT * FROM {tn} WHERE {identifier} =?'.format(tn=signup_list, identifier=sql_id)
    team = cursor.execute(query, (identifier,)).fetchone()[gen_dict[sql_city]]
    return team


def get_role(identifier, cursor):
    """Finds the team of a given id"""
    query = 'SELECT * FROM {tn} WHERE {identifier} =?'.format(tn=signup_list, identifier=sql_id)
    role = cursor.execute(query, (identifier,)).fetchone()[gen_dict[sql_br]]
    return role


def move_selected_contestant(identifier, connection, cursor):
    """Moves dancer, given id, from the selection list to selected list"""
    if identifier is not None:
        query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE id = ?;'.format(tn1=selected_list, tn2=signup_list)
        cursor.execute(query, (identifier,))
        query = 'DELETE FROM {tn} WHERE id = ?'.format(tn=selection_list)
        cursor.execute(query, (identifier,))
        connection.commit()
        logging.info('Selected {} for the NTDS'.format(identifier))


def copy_to_preselection(preselected_people, connection, cursor):
    """Copies a number of dancers to the preselection pool, given a list of people"""
    list_of_ids = [(row[gen_dict[sql_id]],) for row in preselected_people]
    query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
        .format(tn1=preselected_people_list, tn2=selection_list, id=sql_id)
    cursor.executemany(query, list_of_ids)
    connection.commit()


def copy_to_selection_pool(selected_people, connection, cursor):
    """Copies a number of dancers to the selection pool, given a list of people"""
    list_of_ids = [(row[gen_dict[sql_id]],) for row in selected_people]
    query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
        .format(tn1=partners_selection_pool, tn2=selection_list, id=sql_id)
    cursor.executemany(query, list_of_ids)
    connection.commit()


def find_signed_partner(identifier, connection, cursor):
    """"Temp"""
    partner_id = None
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
    dancer = cursor.execute(query, (identifier,)).fetchone()
    if dancer is not None:
        ballroom_partner = dancer[gen_dict[sql_bp]]
        latin_partner = dancer[gen_dict[sql_lp]]
        if isinstance(ballroom_partner, int):
            partner_id = ballroom_partner
        if all([isinstance(latin_partner, int), partner_id is None]):
            partner_id = latin_partner
        # connection.commit()
        if partner_id is None:
            logging.info('{id1} has not signed with a partner (signed)'.format(id1=identifier))
        else:
            logging.info('{id1} and {id2} signed up together (signed)'.format(id1=identifier, id2=partner_id))
    return partner_id


def find_partner(identifier, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=selection_list, id=sql_id)
    dancer = cursor.execute(query, (identifier,)).fetchone()
    if dancer is not None:
        ballroom_level = dancer[gen_dict[sql_bl]]
        latin_level = dancer[gen_dict[sql_ll]]
        ballroom_partner = dancer[gen_dict[sql_bp]]
        latin_partner = dancer[gen_dict[sql_lp]]
        rol = dancer[gen_dict[sql_br]]
        # ballroom_mandatory_date = dancer[gen_dict[sql_bbd]]
        # latin_mandatory_date = dancer[gen_dict[sql_lbd]]
        # team_captain = dancer[gen_dict[sql_tc]]
        team = dancer[gen_dict[sql_city]]
        # Check if the contestant already has signed up with a partner
        if isinstance(ballroom_partner, int):
            partner_id = ballroom_partner
        if all([isinstance(latin_partner, int), partner_id is None]):
            partner_id = latin_partner
        if partner_id is not None:
            logging.info('{id1} and {id2} signed up together'.format(id1=identifier, id2=partner_id))
        # If the contestant is a beginner, and has no partner, find a partner
        query = 'SELECT * FROM {tn} where {bl} = ? AND {ln} = ? AND {bp} = "" AND {lp} = "" AND {br} != ? ' \
                'AND {team} != ?' \
            .format(tn=selection_list, bl=sql_bl, ln=sql_ll, bp=sql_bp, lp=sql_lp, br=sql_br, team=sql_city)
        if all([ballroom_level == beginner, latin_level == beginner, partner_id is None]):
            potential_partners = cursor.execute(query, (ballroom_level, latin_level, rol, team)).fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        # Try to find the best partner for someone that signed alone and has either Breiten or Open level
        if all([ballroom_level == breiten, latin_level == breiten, partner_id is None]):
            potential_partners = cursor.execute(query, (breiten, breiten, rol, team)).fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (breiten, open_class, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (open_class, breiten, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        if all([ballroom_level == breiten, latin_level == open_class, partner_id is None]):
            potential_partners = cursor.execute(query, (breiten, open_class, rol, team)).fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (breiten, breiten, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        if all([ballroom_level == open_class, latin_level == breiten, partner_id is None]):
            potential_partners = cursor.execute(query, (open_class, breiten, rol, team)).fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (breiten, breiten, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        if all([ballroom_level == open_class, latin_level == open_class, partner_id is None]):
            potential_partners = cursor.execute(query, (open_class, open_class, rol, team)).fetchall()
            number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (open_class, breiten, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners == 0:
                potential_partners = cursor.execute(query, (breiten, open_class, rol, team)).fetchall()
                number_of_potential_partners = len(potential_partners)
            if number_of_potential_partners > 0:
                random_num = randint(0, number_of_potential_partners - 1)
                partner_id = potential_partners[random_num][gen_dict[sql_id]]
        # connection.commit()
        if partner_id is None:
            logging.info('Found no match for {id1}'.format(id1=identifier))
        else:
            logging.info('Matched {id1} and {id2} together'.format(id1=identifier, id2=partner_id))
    return partner_id


def find_partner_limited(dancer, selection_pool):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    dancer_id = dancer[gen_dict[sql_id]]
    ballroom_partner = dancer[gen_dict[sql_bp]]
    latin_partner = dancer[gen_dict[sql_lp]]
    rol = dancer[gen_dict[sql_br]]
    team = dancer[gen_dict[sql_city]]
    # Check if the contestant already has signed up with a partner
    if isinstance(ballroom_partner, int):
        partner_id = ballroom_partner
    if all([isinstance(latin_partner, int), partner_id is None]):
        partner_id = latin_partner
    num_of_pot_partners = len(selection_pool)
    if num_of_pot_partners > 0:
        random_order = random.sample(range(0, num_of_pot_partners), num_of_pot_partners)
        if partner_id is None:
            for num in random_order:
                pot_partner = selection_pool[num]
                pot_partner_role = pot_partner[gen_dict[sql_br]]
                pot_partner_partner = pot_partner[gen_dict[sql_bp]]
                pot_partner_team = pot_partner[gen_dict[sql_city]]
                if pot_partner_partner == '':
                    pot_partner_available = True
                else:
                    pot_partner_available = False
                if all([pot_partner_role != rol, pot_partner_available, pot_partner_team != team, partner_id is None]):
                    partner_id = pot_partner[gen_dict[sql_id]]
                    break
    if partner_id is None:
        logging.info('Found no match for {id1} (limited)'.format(id1=dancer_id))
    else:
        logging.info('Matched {id1} and {id2} together (limited)'.format(id1=dancer_id, id2=partner_id))
    return partner_id


def select_guaranteed_beginners(cities_list, cursor):
    """"Temp"""
    guaranteed_beginners = []
    query = 'SELECT * FROM {tn1} WHERE {bl} = ? AND {ln} = ? AND {team} = ?' \
        .format(tn1=selection_list, bl=sql_bl, ln=sql_ll, level=beginner, team=sql_city)
    for city in cities_list:
        city_beginners = cursor.execute(query, (beginner, beginner, city)).fetchall()
        number_of_city_beginners = len(city_beginners)
        if number_of_city_beginners <= max_fixed_beginners:
            guaranteed_beginners.extend(city_beginners)
    return guaranteed_beginners


def select_guaranteed_lions(cities_list, cursor):
    """"Temp"""
    guaranteed_lions = []
    query = get_lions_team_query()
    for city in cities_list:
        city_lions = cursor.execute(query, (city,)).fetchall()
        number_of_city_lions = len(city_lions)
        if number_of_city_lions <= max_fixed_lion_contestants:
            guaranteed_lions.extend(city_lions)
    return guaranteed_lions


def get_lions_team_query():
    """"Temp"""
    query = 'SELECT * FROM {tn1} WHERE {team} = ?' \
        .format(tn1=selection_list, team=sql_city)
    sql_filter = []
    for level in lion_participants:
        sql_filter.append(' ( {bl} = "' + level + '"' + ' OR {ln} = "' + level + '" )')
    query_extension = ' AND (' + ' OR '.join(sql_filter) + ' )'
    query_extension = query_extension.format(bl=sql_bl, ln=sql_ll)
    query += query_extension
    return query


def get_lions_query():
    """"Temp"""
    query = 'SELECT * FROM {tn1} WHERE 1' \
        .format(tn1=selection_list, team=sql_city)
    sql_filter = []
    for level in lion_participants:
        sql_filter.append(' ( {bl} = "' + level + '"' + ' OR {ln} = "' + level + '" )')
    query_extension = ' AND (' + ' OR '.join(sql_filter) + ' )'
    query_extension = query_extension.format(bl=sql_bl, ln=sql_ll)
    query += query_extension
    return query


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


def create_city_beginners_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        # query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ? OR {city_follow} LIKE ?' \
        #     .format(tn=partners_list, city_lead=sql_city_lead, city_follow=sql_city_follow)
        # number_of_city_beginners = len(cursor.execute(query, (city, city)).fetchall())
        query = 'SELECT * FROM {tn1} WHERE {bl} = ? AND {ln} = ? AND {team} = ?' \
            .format(tn1=selection_list, bl=sql_bl, ln=sql_ll, level=beginner, team=sql_city)
        max_city_beginners = len(cursor.execute(query, (beginner, beginner, city)).fetchall())
        if max_city_beginners > max_fixed_beginners:
            max_city_beginners = max_fixed_beginners
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=fixed_beginners_list)
        cursor.execute(query, (city, 0, max_city_beginners))
    connection.commit()


def create_city_lions_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        # query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ? OR {city_follow} LIKE ?' \
        #     .format(tn=partners_list, city_lead=sql_city_lead, city_follow=sql_city_follow)
        # number_of_city_lions = len(cursor.execute(query, (city, city)).fetchall())
        query = 'SELECT * FROM {tn1} WHERE {bl} = ? AND {ln} = ? AND {team} = ?' \
            .format(tn1=selection_list, bl=sql_bl, ln=sql_ll, level=beginner, team=sql_city)
        query = get_lions_team_query()
        max_city_lions = len(cursor.execute(query, (city,)).fetchall())
        if max_city_lions > max_fixed_lion_contestants:
            max_city_lions = max_fixed_lion_contestants
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=fixed_lions_list)
        cursor.execute(query, (city, 0, max_city_lions))
    connection.commit()


def create_pair(first_dancer, second_dancer, connection, cursor):
    """Temp"""
    first_dancer_role = get_role(first_dancer, cursor=cursor)
    if first_dancer_role != 'Lead':
        first_dancer, second_dancer = second_dancer, first_dancer
    if first_dancer is None:
        first_dancer = ''
        first_dancer_team = ''
    else:
        first_dancer_team = get_team(first_dancer, cursor=cursor)
    if second_dancer is None:
        second_dancer = ''
        second_dancer_team = ''
    else:
        second_dancer_team = get_team(second_dancer, cursor=cursor)
    query = 'INSERT INTO {tn} VALUES (?, ?, ?, ?)'.format(tn=partners_list)
    cursor.execute(query, (first_dancer, second_dancer, first_dancer_team, second_dancer_team))
    query = 'INSERT INTO {tn} VALUES (?, ?, ?, ?)'.format(tn=ref_partner_list)
    cursor.execute(query, (first_dancer, second_dancer, first_dancer_team, second_dancer_team))
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=partners_selection_pool, id=sql_id)
    cursor.executemany(query, [(first_dancer,), (second_dancer,)])
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=selection_list, id=sql_id)
    cursor.executemany(query, [(first_dancer,), (second_dancer,)])
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
    tables_to_drop = [signup_list, team_list, selection_list, selected_list, partners_list, partners_selection_pool,
                      preselected_people_list, fixed_beginners_list, fixed_lions_list, ref_partner_list]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        cursor.execute(query)
    # Create new tables
    dancer_list_tables = [signup_list, selection_list, selected_list, partners_selection_pool, preselected_people_list]
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
        query = 'INSERT INTO {} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?);'.format(selection_list)
        for row in city_signup_list:
            curs.execute(query, row)
        total_number_of_contestants += max_row-1
    conn.commit()

    # Select the team captains
    query = 'SELECT * FROM {tn1} WHERE {tc} = "Ja"'.format(tn1=selection_list, tc=sql_tc)
    team_captains = curs.execute(query).fetchall()
    for captain in team_captains:
        captain_id = captain[gen_dict[sql_id]]
        partner_id = find_partner(captain_id, connection=conn, cursor=curs)
        create_pair(captain_id, partner_id, connection=conn, cursor=curs)
        move_selected_contestant(captain_id, connection=conn, cursor=curs)
        move_selected_contestant(partner_id, connection=conn, cursor=curs)
    conn.commit()
    query = drop_table_query.format(partners_list)
    curs.execute(query)
    query = paren_table_query.format(partners_list)
    curs.execute(query)

    # Select the beginners of cities that have less than the maximum guaranteed
    guaranteed_beginners = select_guaranteed_beginners(competing_cities, cursor=curs)
    copy_to_preselection(guaranteed_beginners, connection=conn, cursor=curs)
    create_city_beginners_list(competing_cities, connection=conn, cursor=curs)
    query = 'SELECT * FROM {tn1} WHERE {bl} = "{level}" AND {ln} = "{level}"' \
        .format(tn1=selection_list, bl=sql_bl, ln=sql_ll, level=beginner)
    all_beginners = curs.execute(query).fetchall()
    copy_to_selection_pool(all_beginners, connection=conn, cursor=curs)
    for beg in guaranteed_beginners:
        beginner_id = beg[gen_dict[sql_id]]
        query = ' SELECT * FROM {tn} WHERE {id} = ?'.format(tn=partners_selection_pool, id=sql_id)
        beginner_available = curs.execute(query, (beginner_id,)).fetchone()
        if beginner_available is not None:
            partner_id = None
            query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_beginners_list, num=sql_num_con)
            ordered_cities = curs.execute(query).fetchall()
            for city in ordered_cities:
                city = city[city_dict[sql_city]]
                if partner_id is None:
                    query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
                    sel_pool = curs.execute(query, (city,)).fetchall()
                    partner_id = find_partner_limited(beg, selection_pool=sel_pool)
                    if partner_id is not None:
                        create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                        move_selected_contestant(beginner_id, connection=conn, cursor=curs)
                        move_selected_contestant(partner_id, connection=conn, cursor=curs)
                        update_city_beginners(competing_cities, connection=conn, cursor=curs)
    # Get rest of beginners
    query = 'SELECT sum({max_beg})-sum({num}) FROM {tn}'\
        .format(tn=fixed_beginners_list, max_beg=sql_max_con, num=sql_num_con)
    max_iterations = curs.execute(query).fetchone()[0]
    max_iterations = int(round((max_iterations+1)/2))+1
    for iteration in range(max_iterations):
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_beginners_list, num=sql_num_con)
        ordered_cities = curs.execute(query).fetchall()
        selected_city = None
        for city in ordered_cities:
            if selected_city is None:
                number_of_selected_city_beginners = city[city_dict[sql_num_con]]
                max_number_of_city_beginners = city[city_dict[sql_max_con]]
                if number_of_selected_city_beginners < max_number_of_city_beginners:
                    selected_city = city[city_dict[sql_city]]
        if selected_city is not None:
            query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
            selected_city_beginners = curs.execute(query, (selected_city,)).fetchall()
            number_of_city_beginners = len(selected_city_beginners)
            if number_of_city_beginners > 0:
                random_order = random.sample(range(0, number_of_city_beginners), number_of_city_beginners)
                partner_id = None
                for order_city in ordered_cities:
                    if partner_id is None:
                        order_city = order_city[city_dict[sql_city]]
                        query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
                        order_city_beginners = curs.execute(query, (order_city,)).fetchall()
                        available_beginners = len(order_city_beginners)
                        if order_city != selected_city and available_beginners > 0:
                            query = 'SELECT * FROM {tn} WHERE {team} = ?'\
                                .format(tn=partners_selection_pool, team=sql_city)
                            sel_pool = curs.execute(query, (order_city,)).fetchall()
                            for num in random_order:
                                if partner_id is None:
                                    beg = selected_city_beginners[num]
                                    beginner_id = beg[gen_dict[sql_id]]
                                    query = ' SELECT * FROM {tn} WHERE {id} = ?'\
                                        .format(tn=partners_selection_pool,id=sql_id)
                                    beginner_available = curs.execute(query, (beginner_id,)).fetchone()
                                    if beginner_available is not None:
                                        partner_id = find_partner_limited(beg, sel_pool)
                                        if partner_id is not None:
                                            create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                                            move_selected_contestant(beginner_id, connection=conn, cursor=curs)
                                            move_selected_contestant(partner_id, connection=conn, cursor=curs)
                                            update_city_beginners(competing_cities, connection=conn, cursor=curs)

    if testing:
        query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=fixed_beginners_list, city=sql_city)
        ordered_cities = curs.execute(query).fetchall()
        query = 'CREATE TABLE IF NOT EXISTS "beginners" ' \
                '(id INTEGER PRIMARY KEY AUTOINCREMENT, {c1} INT, {c2} INT, {c3} INT, {c4} INT)' \
            .format(c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0], c4=ordered_cities[3][0])
        curs.execute(query)
        query = 'INSERT INTO "beginners" ({c1},{c2},{c3},{c4}) VALUES (?,?,?,?)'\
            .format(c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0], c4=ordered_cities[3][0])
        curs.execute(query, (ordered_cities[0][1], ordered_cities[1][1], ordered_cities[2][1], ordered_cities[3][1]))
        for city in ordered_cities:
            overview = 'Number of selected beginners from {city} is: {number}'.format(city=city[0], number=city[1])
            print(overview)
        conn.commit()

    tables_to_drop = [partners_list, partners_selection_pool, preselected_people_list]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        curs.execute(query)
    conn.commit()
    query = paren_table_query.format(partners_list)
    curs.execute(query)
    query = dancers_list_query.format(partners_selection_pool)
    curs.execute(query)
    query = dancers_list_query.format(preselected_people_list)
    curs.execute(query)
    conn.commit()

    # Get lion contestants of cities that have less than the maximum guaranteed
    guaranteed_lions = select_guaranteed_lions(competing_cities, cursor=curs)
    copy_to_preselection(guaranteed_lions, connection=conn, cursor=curs)
    create_city_lions_list(competing_cities, connection=conn, cursor=curs)
    query = get_lions_query()
    all_lions = curs.execute(query).fetchall()
    copy_to_selection_pool(all_lions, connection=conn, cursor=curs)
    for lion in guaranteed_lions:
        lion_id = lion[gen_dict[sql_id]]
        query = ' SELECT * FROM {tn} WHERE {id} = ?'.format(tn=selected_list, id=sql_id)
        lion_available = curs.execute(query, (lion_id,)).fetchone()
        if lion_available is None:
            partner_id = None
            query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_lions_list, num=sql_num_con)
            ordered_cities = curs.execute(query).fetchall()
            for city in ordered_cities:
                city = city[city_dict[sql_city]]
                if partner_id is None:
                    query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
                    sel_pool = curs.execute(query, (city,)).fetchall()
                    partner_id = find_partner_limited(lion, selection_pool=sel_pool)
                    if partner_id is not None:
                        create_pair(lion_id, partner_id, connection=conn, cursor=curs)
                        move_selected_contestant(lion_id, connection=conn, cursor=curs)
                        move_selected_contestant(partner_id, connection=conn, cursor=curs)
                        update_city_lions(competing_cities, connection=conn, cursor=curs)
    # Get rest of lions
    query = 'SELECT sum({max_lion})-sum({num}) FROM {tn}' \
        .format(tn=fixed_lions_list, max_lion=sql_max_con, num=sql_num_con)
    max_iterations = curs.execute(query).fetchone()[0]
    max_iterations = int(round((max_iterations + 1) / 2)) + 1
    for iteration in range(max_iterations):
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=fixed_lions_list, num=sql_num_con)
        ordered_cities = curs.execute(query).fetchall()
        selected_city = None
        for city in ordered_cities:
            if selected_city is None:
                number_of_selected_city_lions = city[city_dict[sql_num_con]]
                max_number_of_city_lions = city[city_dict[sql_max_con]]
                if number_of_selected_city_lions < max_number_of_city_lions:
                    selected_city = city[city_dict[sql_city]]
        if selected_city is not None:
            query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
            city_lions = curs.execute(query, (selected_city,)).fetchall()
            number_of_city_lions = len(city_lions)
            if number_of_city_lions > 0:
                random_order = random.sample(range(0, number_of_city_lions), number_of_city_lions)
                partner_id = None
                for order_city in ordered_cities:
                    if partner_id is None:
                        order_city = order_city[city_dict[sql_city]]
                        query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_city)
                        order_city_lions = curs.execute(query, (order_city,)).fetchall()
                        available_lions = len(order_city_lions)
                        if order_city != selected_city and available_lions > 0:
                            query = 'SELECT * FROM {tn} WHERE {team} = ?' \
                                .format(tn=partners_selection_pool, team=sql_city)
                            sel_pool = curs.execute(query, (order_city,)).fetchall()
                            for num in random_order:
                                if partner_id is None:
                                    lion = city_lions[num]
                                    lion_id = lion[gen_dict[sql_id]]
                                    query = ' SELECT * FROM {tn} WHERE {id} = ?' \
                                        .format(tn=partners_selection_pool, id=sql_id)
                                    lion_available = curs.execute(query, (lion_id,)).fetchone()
                                    if lion_available is not None:
                                        partner_id = find_partner_limited(lion, sel_pool)
                                        if partner_id is not None:
                                            create_pair(lion_id, partner_id, connection=conn, cursor=curs)
                                            move_selected_contestant(lion_id, connection=conn, cursor=curs)
                                            move_selected_contestant(partner_id, connection=conn, cursor=curs)
                                            update_city_lions(competing_cities, connection=conn, cursor=curs)

    if testing:
        query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=fixed_lions_list, city=sql_city)
        ordered_cities = curs.execute(query).fetchall()
        query = 'CREATE TABLE IF NOT EXISTS "lions" ' \
                '(id INTEGER PRIMARY KEY AUTOINCREMENT, {c1} INT, {c2} INT, {c3} INT, {c4} INT)' \
            .format(c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0], c4=ordered_cities[3][0])
        curs.execute(query)
        query = 'INSERT INTO "lions" ({c1},{c2},{c3},{c4}) VALUES (?,?,?,?)' \
            .format(c1=ordered_cities[0][0], c2=ordered_cities[1][0], c3=ordered_cities[2][0], c4=ordered_cities[3][0])
        curs.execute(query, (ordered_cities[0][1], ordered_cities[1][1], ordered_cities[2][1], ordered_cities[3][1]))
        for city in ordered_cities:
            overview = 'Number of selected lions from {city} is: {number}'.format(
                city=city[0], number=city[1])
            print(overview)
        conn.commit()

    query = 'SELECT * FROM {tn}'.format(tn=selected_list, num=sql_num_con)
    selected_dancers = curs.execute(query).fetchall()
    number_of_selected_dancers = len(selected_dancers)
    # max_iterations = max_contestants - buffer_for_selection - number_of_selected_dancers
    # max_iterations = int(round((max_iterations + 1) / 2))
    query = 'SELECT * FROM {tn} ORDER BY RANDOM()'.format(tn=selection_list)
    available_dancers = curs.execute(query).fetchall()
    number_of_available_dancers = len(available_dancers)
    # if number_of_available_dancers < int(round((number_of_available_dancers + 1) / 2)):
    #     max_iterations = int(round((number_of_available_dancers + 1) / 2))
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
            all_dancers = curs.execute(query).fetchall()
            number_of_selected_dancers = len(all_dancers) + buffer_for_selection
            if number_of_selected_dancers >= max_contestants:
                break
    conn.commit()
    query = drop_table_query.format(partners_list)
    curs.execute(query)
    query = paren_table_query.format(partners_list)
    curs.execute(query)

    # Close cursor and connection
    print("--- Done in %.3f seconds ---" % (time.time() - start_time))
    curs.close()
    conn.close()

if __name__ == '__main__':
    for i in range(1):
        main()
