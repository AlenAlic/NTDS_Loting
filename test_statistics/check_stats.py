import os.path
import configparser
from check_signupsheet import column
import statistics

file_key = {'name': 'Statistics_and_Notes_4.ini', 'path': ''}
file_key['path'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_key['path'])
file_key['path'] = os.path.join(file_key['path'], file_key['name'])

values = list()
keys = None

config_parser = configparser.ConfigParser()
if os.path.isfile(path=file_key['path']):
    config_parser.read(file_key['path'])
    for sec in config_parser.sections():
        values.append(list(map(int, list(dict(config_parser.items(sec)).values()))))
    keys = dict(config_parser.items('1'))

means = dict()
stdevs = dict()
end_list = list()

if keys is not None:
    for index, key in enumerate(keys.keys()):
        means[key] = sum(column(values, index))/1000
    for index, key in enumerate(keys.keys()):
        stdevs[key] = statistics.pstdev(column(values, index))
    for index, key in enumerate(means.keys()):
        end_list.append([str(key), str(means[key]), str(stdevs[key])])
    for item in end_list:
        print(item)

print('')
