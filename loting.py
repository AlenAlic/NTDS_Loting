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
ref_partner_list = 'reference_partner_list'
preselected_people_list = 'preselected_people'
city_list = 'city_list'

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

# SQL Table column names and dictionary for city list
sql_beg = 'number_of_beginners'
sql_max_beg = 'max_beginners'
city_dict = {sql_city: 0, sql_beg: 1, sql_max_beg: 2}

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
city_list_query = 'CREATE TABLE {tn} ({city} TEXT PRIMARY KEY, {beg} INT, {max_beg} INT);' \
    .format(tn={}, city=sql_city, beg=sql_beg, max_beg=sql_max_beg)


def get_team(identifier, connection, cursor):
    """Finds the team of a given id"""
    query = 'SELECT * FROM {tn} WHERE {identifier} =?'.format(tn=signup_list, identifier=sql_id)
    team = cursor.execute(query, (identifier,)).fetchone()[gen_dict[sql_team]]
    connection.commit()
    return team


def get_role(identifier, connection, cursor):
    """Finds the team of a given id"""
    query = 'SELECT * FROM {tn} WHERE {identifier} =?'.format(tn=signup_list, identifier=sql_id)
    role = cursor.execute(query, (identifier,)).fetchone()[gen_dict[sql_br]]
    connection.commit()
    return role


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
    connection.commit()


def find_signed_partner(dancer):
    """Finds the partner a dancer signed up with (if it exists), given id"""
    partner_id = None
    dancer_id = dancer[gen_dict[sql_id]]
    ballroom_partner = dancer[gen_dict[sql_bp]]
    latin_partner = dancer[gen_dict[sql_lp]]
    if isinstance(ballroom_partner, int):
        partner_id = ballroom_partner
    if all([isinstance(latin_partner, int), partner_id is None]):
        partner_id = latin_partner
    if partner_id is None:
        logging.info('Found no match for {id1}'.format(id1=dancer_id))
    else:
        logging.info('Matched {id1} and {id2} together'.format(id1=dancer_id, id2=partner_id))
    return partner_id


def find_partner(identifier, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    query = 'SELECT * from {tn} WHERE {id} =?'.format(tn=signup_list, id=sql_id)
    dancer = cursor.execute(query, (identifier,)).fetchone()
    ballroom_level = dancer[gen_dict[sql_bn]]
    latin_level = dancer[gen_dict[sql_ln]]
    ballroom_partner = dancer[gen_dict[sql_bp]]
    latin_partner = dancer[gen_dict[sql_lp]]
    rol = dancer[gen_dict[sql_br]]
    # ballroom_mandatory_date = dancer[gen_dict[sql_bbd]]
    # latin_mandatory_date = dancer[gen_dict[sql_lbd]]
    # team_captain = dancer[gen_dict[sql_tc]]
    team = dancer[gen_dict[sql_team]]
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


def find_beginner_partner(dancer, selection_pool, connection, cursor):
    """Finds the/a partner for a dancer, given id"""
    partner_id = None
    dancer_id = dancer[gen_dict[sql_id]]
    latin_level = dancer[gen_dict[sql_ln]]
    ballroom_partner = dancer[gen_dict[sql_bp]]
    latin_partner = dancer[gen_dict[sql_lp]]
    rol = dancer[gen_dict[sql_br]]
    team = dancer[gen_dict[sql_team]]
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
                if pot_partner_partner == '':
                    pot_partner_available = True
                else:
                    pot_partner_available = False
                if all([pot_partner_role != rol, pot_partner_available, partner_id is None]):
                    partner_id = pot_partner[gen_dict[sql_id]]
                    break
    if partner_id is None:
        logging.info('Found no match for {id1}'.format(id1=dancer_id))
    else:
        logging.info('Matched {id1} and {id2} together'.format(id1=dancer_id, id2=partner_id))
    return partner_id


def select_guaranteed_beginners(cities_list, cursor):
    """"Temp"""
    guaranteed_beginners = []
    for city in cities_list:
        query = 'SELECT * FROM {tn1} WHERE {bn} = ? AND {ln} = ? AND {team} = ?'\
            .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team)
        city_beginners = cursor.execute(query, (beginner, beginner, city)).fetchall()
        number_of_city_beginners = len(city_beginners)
        if number_of_city_beginners <= max_fixed_beginners:
            guaranteed_beginners.extend(city_beginners)
    return guaranteed_beginners


def update_city_beginners_completeness(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ?' \
            .format(tn=partners_list, city_lead=sql_city_lead)
        number_of_city_beginners = len(cursor.execute(query, (city,)).fetchall())
        query = 'SELECT * FROM {tn} WHERE {city_follow} LIKE ?' \
            .format(tn=partners_list, city_follow=sql_city_follow)
        number_of_city_beginners += len(cursor.execute(query, (city,)).fetchall())
        query = 'UPDATE {tn} SET {num} = ? WHERE {city} = ?' \
            .format(tn=city_list, num=sql_beg, city=sql_city)
        cursor.execute(query, (number_of_city_beginners, city))
    connection.commit()


def create_city_list(cities_list, connection, cursor):
    """"Temp"""
    for city in cities_list:
        query = 'SELECT * FROM {tn} WHERE {city_lead} LIKE ? OR {city_follow} LIKE ?' \
            .format(tn=partners_list, city_lead=sql_city_lead, city_follow=sql_city_follow)
        number_of_city_beginners = len(cursor.execute(query, (city, city)).fetchall())
        query = 'SELECT * FROM {tn1} WHERE {bn} = ? AND {ln} = ? AND {team} = ?' \
            .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner, team=sql_team)
        max_city_beginners = len(cursor.execute(query, (beginner, beginner, city)).fetchall())
        if max_city_beginners > max_fixed_beginners:
            max_city_beginners = max_fixed_beginners
        query = 'INSERT INTO {tn} VALUES (?, ?, ?)'.format(tn=city_list)
        cursor.execute(query, (city, number_of_city_beginners, max_city_beginners))
    connection.commit()


def create_pair(first_dancer, second_dancer, connection, cursor):
    """Temp"""
    first_dancer_role = get_role(first_dancer, connection=connection, cursor=cursor)
    if first_dancer_role != 'Lead':
        first_dancer, second_dancer = second_dancer, first_dancer
    first_dancer_team = get_team(first_dancer, connection=connection, cursor=cursor)
    second_dancer_team = get_team(second_dancer, connection=connection, cursor=cursor)
    query = 'INSERT INTO {tn} VALUES (?, ?, ?, ?)'.format(tn=partners_list)
    cursor.execute(query, (first_dancer, second_dancer, first_dancer_team, second_dancer_team))
    query = 'DELETE FROM {tn} WHERE {id} = ?'.format(tn=partners_selection_pool, id=sql_id)
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
    tables_to_drop = [signup_list, selected_list, team_list, partners_list, partners_selection_pool,
                      preselected_people_list, city_list, ref_partner_list]
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
    query = paren_table_query.format(ref_partner_list)
    cursor.execute(query)
    query = city_list_query.format(city_list)
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
    guaranteed_beginners = select_guaranteed_beginners(competing_cities, cursor=curs)
    copy_to_preselection(guaranteed_beginners, connection=conn, cursor=curs)
    create_city_list(competing_cities, connection=conn, cursor=curs)
    query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}"' \
        .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner)
    all_beginners = curs.execute(query).fetchall()
    copy_to_selection_pool(all_beginners, connection=conn, cursor=curs)
    for beg in guaranteed_beginners:
        beginner_id = beg[gen_dict[sql_id]]
        beg_city = beg[gen_dict[sql_team]]
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=city_list, num=sql_beg)
        ordered_cities = curs.execute(query).fetchall()
        partner_id = None
        for city in ordered_cities:
            city = city[city_dict[sql_city]]
            if city != beg_city:
                query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_team)
                sel_pool = curs.execute(query, (city,)).fetchall()
                if partner_id is None:
                    partner_id = find_beginner_partner(beg, selection_pool=sel_pool, connection=conn, cursor=curs)
                    if partner_id is not None:
                        create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                        update_city_beginners_completeness(competing_cities, connection=conn, cursor=curs)
    # Get rest of beginners
    query = 'SELECT sum({max_beg})-sum({num}) from {tn}'.format(tn=city_list, max_beg=sql_max_beg, num=sql_beg)
    max_iterations = curs.execute(query).fetchone()[0]
    max_iterations = int(round((max_iterations+1)/2))+1
    for iteration in range(max_iterations):
        query = 'SELECT * FROM {tn} ORDER BY {num}, RANDOM()'.format(tn=city_list, num=sql_beg)
        ordered_cities = curs.execute(query).fetchall()
        selected_city = None
        for city in ordered_cities:
            number_of_selected_city_beginners = city[city_dict[sql_beg]]
            max_number_of_city_beginners = city[city_dict[sql_max_beg]]
            if number_of_selected_city_beginners < max_number_of_city_beginners and selected_city is None:
                selected_city = city[city_dict[sql_city]]
        if selected_city is not None:
            query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_team)
            city_beginners = curs.execute(query, (selected_city,)).fetchall()
            number_of_city_beginners = len(city_beginners)
            if number_of_city_beginners > 0:
                random_order = random.sample(range(0, number_of_city_beginners), number_of_city_beginners)
                partner_id = None
                for order_city in ordered_cities:
                    if partner_id is None:
                        order_city = order_city[city_dict[sql_city]]
                        if order_city != selected_city:
                            query = 'SELECT * FROM {tn} WHERE {team} = ?'.format(tn=partners_selection_pool, team=sql_team)
                            sel_pool = curs.execute(query, (order_city,)).fetchall()
                            for num in random_order:
                                beg = city_beginners[num]
                                beginner_id = beg[gen_dict[sql_id]]
                                if partner_id is None:
                                    partner_id = find_beginner_partner(beg, sel_pool, connection=conn, cursor=curs)
                                    if partner_id is not None:
                                        create_pair(beginner_id, partner_id, connection=conn, cursor=curs)
                                        update_city_beginners_completeness(competing_cities, connection=conn, cursor=curs)

    if testing:
        query = 'SELECT * FROM {tn} ORDER BY {city}'.format(tn=city_list, city=sql_city)
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

    # Move selected pairs to selected people
    # Keep track of made pairs

    # Close cursor and connection
    curs.close()
    conn.close()
    print("--- Done in %s seconds ---" % (time.time() - start_time))

if __name__ == '__main__':
    for i in range(1):
        main()
