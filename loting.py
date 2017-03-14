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
p_dict = {sql_lead: 0, sql_follow: 1}

# General query formats
drop_table_query = 'DROP TABLE IF EXISTS {};'
dancers_list_query = 'CREATE TABLE {tn} ({id} INT PRIMARY KEY, {name} TEXT, {email} TEXT, {bn} TEXT, {ln} TEXT,' \
                     ' {bp} INT, {lp} INT, {br} TEXT, {lr} TEXT, {bbd} TEXT, {lbd} TEXT, {tc} TEXT, {team} TEXT);'\
                     .format(tn={}, id=sql_id, name=sql_name, email=sql_email, bn=sql_bn, ln=sql_ln, bp=sql_bp,
                             lp=sql_lp, br=sql_br, lr=sql_lr, bbd=sql_bbd, lbd=sql_lbd, tc=sql_tc, team=sql_team)
team_list_query = 'CREATE TABLE {tn} ({team} TEXT PRIMARY KEY, {city} TEXT, {signup_list} TEXT);' \
    .format(tn={}, team=sql_team, city=sql_city, signup_list=sql_sl)
paren_table_query = 'CREATE TABLE {tn} ({lead} INT PRIMARY KEY, {follow} INT);' \
    .format(tn={}, lead=sql_lead, follow=sql_follow)


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


def copy_to_selection_pool(list_of_ids, connection, cursor):
    """Copies a number of dancers to the selection pool, given a list of id's"""
    query = 'INSERT INTO {tn1} SELECT * FROM {tn2} WHERE {id} = ?;'\
        .format(tn1=partners_selection_pool, tn2=signup_list, id=sql_id)
    cursor.executemany(query, list_of_ids)
    connection.commit()


def find_beginner_from_own_team(identifier, connection, cursor):
    """"Finds the/a partner for the first round selection of the beginners"""
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


def create_paar(lead, follow, connection, cursor):
    """Temp"""
    query = 'INSERT INTO {tn} VALUES(?, ?)'.format(tn=partners_list)
    cursor.execute(query, (lead, follow))
    query = 'DELETE FROM {tn} WHERE id = ?'.format(tn=partners_selection_pool)
    cursor.executemany(query, [(lead,), (follow,)])
    connection.commit()


def no_partner_found(identifier, connection, cursor):
    """Temp"""
    query = 'DELETE FROM {tn} WHERE id = ?'.format(tn=partners_selection_pool)
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
    tables_to_drop = [signup_list, selected_list, team_list, partners_list, partners_selection_pool]
    for item in tables_to_drop:
        query = drop_table_query.format(item)
        cursor.execute(query)
    # Create new tables
    dancer_list_tables = [signup_list, selected_list, partners_selection_pool]
    for item in dancer_list_tables:
        query = dancers_list_query.format(item)
        cursor.execute(query)
    query = team_list_query.format(team_list)
    cursor.execute(query)
    query = paren_table_query.format(partners_list)
    cursor.execute(query)
    connection.commit()


def create_competing_cities_list():
    """"Creates list of all competing cities"""
    # Create Workbook object and a Worksheet from it
    wb = openpyxl.load_workbook(participating_teams)
    ws = wb.worksheets[0]
    # Empty 2d list that will contain the signup sheet filenames for each of the teams
    competing_cities_list = []
    # Fill up list with competing cities
    for i in range(2, ws.max_row + 1):
        competing_cities_list \
            .append([ws.cell(row=i, column=1).value, ws.cell(row=i, column=2).value, ws.cell(row=i, column=3).value])
    return competing_cities_list


def main():
    start_time = time.time()
    # Connect to database and create a cursor
    conn = sql.connect(database_name)
    curs = conn.cursor()
    # Create SQL tables
    create_tables(connection=conn, cursor=curs)
    # Create competing cities list
    competing_cities_list = create_competing_cities_list()
    number_of_competing_cities = len(competing_cities_list)
    # Get maximum number of columns
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    max_col = max_rc('col', ws)
    # Copy the signup list from every team into the SQL database
    total_signup_list = list()
    total_number_of_contestants = 0
    for i in range(0, number_of_competing_cities):
        city = competing_cities_list[i][team_dict[sql_city]]
        # Get maximum number of rows and extract signup list
        wb = openpyxl.load_workbook(competing_cities_list[i][team_dict[sql_sl]])
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

    # Select the team captains from each team
    for competing_city in competing_cities_list:
        city = competing_city[team_dict[sql_city]]
        query = 'SELECT * FROM {tn1} WHERE {tc} = "Ja" AND {team} = "{city}"'\
            .format(tn1=signup_list, city=city, tc=sql_tc, team=sql_team)
        team_captains = curs.execute(query).fetchall()
        for captain in team_captains:
            captain_id = captain[gen_dict[sql_id]]
            partner_id = find_partner(captain_id, connection=conn, cursor=curs)
            move_selected_contestant(captain_id, connection=conn, cursor=curs)
            move_selected_contestant(partner_id, connection=conn, cursor=curs)
    conn.commit()

    # get beginners
    for i in range(0, number_of_competing_cities):
        # Reset the tables used to select partners
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
        city = competing_cities_list[i][team_dict[sql_city]]
        query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {ln} = "{level}"' \
            .format(tn1=signup_list, bn=sql_bn, ln=sql_ln, level=beginner)
        all_beginners = curs.execute(query).fetchall()
        list_of_ids = [(row[gen_dict[sql_id]],) for row in all_beginners]
        copy_to_selection_pool(list_of_ids, connection=conn, cursor=curs)
        # Temp
        query = 'SELECT min({id}) FROM {tn1} WHERE {team} = "{city}"' \
            .format(id=sql_id, tn1=partners_selection_pool, team=sql_team, city=city)
        min_id = curs.execute(query).fetchone()[gen_dict[sql_id]]
        # Get beginners of city
        query = 'SELECT * FROM {tn1} WHERE {team} = "{city}"' \
            .format(tn1=partners_selection_pool, team=sql_team, city=city)
        beginners_city = curs.execute(query).fetchall()
        number_of_beginners = len(beginners_city)
        # min_id = 0
        while number_of_beginners > 0:
            random_int = randint(0, number_of_beginners - 1)
            beginner_id = beginners_city[random_int][gen_dict[sql_id]]
            partner_id = find_partner(beginner_id, connection=conn, cursor=curs)
            if partner_id is not None:
                if beginners_city[random_int][gen_dict[sql_br]] == 'Lead':
                    create_paar(lead=beginner_id, follow=partner_id, connection=conn, cursor=curs)
                else:
                    create_paar(lead=partner_id, follow=beginner_id, connection=conn, cursor=curs)
                query = 'SELECT * FROM {tn} WHERE {team} = "{city}"'\
                    .format(tn=partners_selection_pool, team=sql_team, city=city)
            else:
                no_partner_found(beginner_id, connection=conn, cursor=curs)
            beginners_city = curs.execute(query).fetchall()
            number_of_beginners = len(beginners_city)
        query = 'SELECT * FROM {tn} WHERE {lead} >= ? AND {follow} >= ? ORDER BY {lead}'\
            .format(tn=partners_list, lead=sql_lead, follow=sql_follow)
        paren = curs.execute(query, [(min_id), (min_id)]).fetchall()
        aantal_paren = len(paren)
        query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"' \
            .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=city)
        number_of_selected_beginners = len(curs.execute(query).fetchall())
        if aantal_paren + number_of_selected_beginners < max_fixed_beginners:
            print('dummy')
            teams_paren = []
            for row in paren:
                teams_paren.append(get_team(row[0], connection=conn, cursor=curs))
                teams_paren.append(get_team(row[1], connection=conn, cursor=curs))
            for contestant in teams_paren:
                move_selected_contestant(contestant, connection=conn, cursor=curs)
            query = 'SELECT * FROM {tn} WHERE {lead} >= ? AND {follow} >= ? ORDER BY {lead}' \
                .format(tn=partners_list, lead=sql_lead, follow=sql_follow)
            paren = curs.execute(query, [(min_id), (min_id)]).fetchall()
            aantal_paren = len(paren)
            query = 'SELECT * FROM {tn1} WHERE {bn} = "{level}" AND {team} = "{city}"' \
                .format(tn1=selected_list, bn=sql_bn, level=beginner, team=sql_team, city=city)
            number_of_selected_beginners = len(curs.execute(query).fetchall())
        if aantal_paren <= max_fixed_beginners:
            random_numbers = list(range(aantal_paren))
        if aantal_paren > max_fixed_beginners:
            random_numbers = random.sample(range(0, aantal_paren-1), max_fixed_beginners)
        geselecteerde_paren = [paren[random_numbers[i]] for i in range(0, len(random_numbers))]
        teams_paren = []
        for j in range(len(random_numbers)):
            teams_paren.append(get_team(geselecteerde_paren[j][0], connection=conn, cursor=curs))
            teams_paren.append(get_team(geselecteerde_paren[j][1], connection=conn, cursor=curs))
            if teams_paren.count(city) + number_of_selected_beginners >= max_fixed_beginners:
                break
        geselecteerde_paren = geselecteerde_paren[0:int(len(teams_paren) / 2)]
        geselecteerde_paren = [j for k in geselecteerde_paren for j in k]
        for j in range(len(geselecteerde_paren)):
            move_selected_contestant(geselecteerde_paren[j], connection=conn, cursor=curs)

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

    # Close cursor and connection
    curs.close()
    conn.close()
    print("--- %s seconds ---" % (time.time() - start_time))
    print('Done!')

if __name__ == '__main__':
    main()
