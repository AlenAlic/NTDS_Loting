import os.path
import configparser
from check_signupsheet import column
import statistics
from tabulate import tabulate

file_key = {'name': 'NTDS_Status_1513639050.ini', 'path': ''}
file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['path'])
file_key['path'] = os.path.join(file_key['path'], file_key['name'])

# number_of_runs = 1000

values = list()
keys = None

config_parser = configparser.ConfigParser()
if os.path.isfile(path=file_key['path']):
    config_parser.read(file_key['path'])
    for sec in config_parser.sections():
        values.append(list(map(int, list(dict(config_parser.items(sec)).values()))))
    keys = dict(config_parser.items('1'))

number_of_runs = int(sec)
means = dict()
stdevs = dict()
minimal_values = dict()
maximal_values = dict()
end_list = list()
table_title = ['item', 'mean', 'sigma', 'min', 'max']

if keys is not None:
    for index, key in enumerate(keys.keys()):
        means[key] = int(round(sum(column(values, index))/number_of_runs, 0))
    for index, key in enumerate(keys.keys()):
        stdevs[key] = round(statistics.pstdev(column(values, index)), 1)
    for index, key in enumerate(keys.keys()):
        minimal_values[key] = min(column(values, index))
    for index, key in enumerate(keys.keys()):
        maximal_values[key] = max(column(values, index))
    for index, key in enumerate(means.keys()):
        end_list.append([str(key), str(means[key]), str(stdevs[key]),
                         str(minimal_values[key]), str(maximal_values[key])])
    # for item in end_list:
    #     print(item)

print(tabulate(end_list, headers=table_title))


print('')
