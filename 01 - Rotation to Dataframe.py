'''
Instructions for running this notebook:
1. Upload the files into the folder
2. Copy the file name
3. Edit "document_name" variable down below
4. Edit 'output_name" variable down below

'''

# If this is your first time running this notebook, run "pip install docx2txt"
#pip install docx2txt

import docx2txt
import re
import pandas as pd
from datetime import datetime

# Edit this
document_name = r'2022-07-09_Rubber Tire draft Weekend Rotation.docx'
# Edit this
output_name = r'2022-07-09_Rubber Tire draft WEND Rotation_cleaned rotation'

def delimit_by_length(textstring, n):
    split_string = []
    for i in range(0, len(textstring), n):
        split_string.append(textstring[i:i + n])
    return split_string

# Import text from document
rotation_text = docx2txt.process(document_name)

# Split by new line
rotation_lines = [l for l in rotation_text.splitlines() if l != '']

# Find all the routes in this spreadsheet
route_list = []

# This regular expression finds the words after 'LINE '
pattern = re.compile(r'(?<=LINE ).*(?=\b)')
pattern2 = re.compile(r'(?<=ROUTE ).*(?=\b)')
pattern3 = re.compile(r'(?<=  \b).*(?=\b)')

for matchline in [line for line in rotation_lines if line[:7] == 'SERVICE']:
    result = pattern.findall(matchline)
    if len(result) == 0:
        result = pattern2.findall(matchline)
    if len(result) == 0:
        result = pattern3.findall(matchline)
    if len(result) != 0:
        if result[0] not in route_list:
            route_list.append(result[0])

# Find the corresponding division
division_dict = {}

# This regular expression finds the words after 'DIVISION: '
pattern3 = re.compile(r'(?<=DIVISION.: )(?:\S|\s(?!\s{2,}))*')

route_index = 0
x = 0
for i, j in enumerate(rotation_lines):
    for route in route_list:
        if route in j:
            matchline = rotation_lines[i + 1]
            result = pattern3.findall(matchline)
            if len(result) != 0 and route not in division_dict.keys():
                   division_dict[route] = result[0]

def scanfordirection(inputstr):
    pattern4 = re.compile(r'\d\b\s*\b(\S*BOUND)')
    result = pattern4.findall(inputstr)
    if len(result) != 0:
        return result[0]
    else:
        return ''

# Find where each route rotation begins
last_matched_route = []
slicing_list = []

for i, j in enumerate(rotation_lines):
    for route in route_list:
        if route in j:
            direction = scanfordirection(rotation_lines[i+1])
            routedir = route + ' ' + direction
            if routedir not in last_matched_route:
                last_matched_route.append(routedir)
                slicing_list.append(i)

slicing_list.append(len(rotation_lines))
for r in last_matched_route:
    print(r)

# Group rotations by route
slicingtuples = [(slicing_list[i],slicing_list[i+1]) for i in range(0,len(slicing_list)-1)]

rotation_dict = {}
for x,(y,z) in enumerate(slicingtuples):
    sliced_list = rotation_lines[y:z]
    rotation_dict[last_matched_route[x]] = sliced_list

# Drop all non-timepoint header rows
lines_to_drop = ['PROCESSED:','SIGN UP  :','SERVICE  :','DIVISION :','DIVISIONS:','SCENARIO :','SCENARIOS:','==========','          ']

for route in last_matched_route:
    rotation_dict[route] = [d for d in rotation_dict[route] if d[:10] not in lines_to_drop]

# Combine first two timepoint label rows into one and create a header row
rotation_dict_headers = {}
for k in rotation_dict.keys():
    row1 = rotation_dict[k][0]
    header1 = [row1[1:3], row1[4], row1[6], row1[8:11], 'Run_Start', row1[17:21], row1[23:26]] + delimit_by_length(row1[28:],5)
    row2 = rotation_dict[k][1]
    header2 = [row2[:3], '', row2[6], '', '', row2[17:21], row2[22:26]] + delimit_by_length(row2[28:],5)
    rotation_dict_headers[k] = [x + y for (x,y) in zip(header1, header2)]

# Drop the old timepoint label rows
lines_to_drop2 = [' LN ',' NO ']

for route in last_matched_route:
    rotation_dict[route] = [d for d in rotation_dict[route] if d[:4] not in lines_to_drop2]

# Parsing function
def parserotation(text):
        if len(text) > 25:
            return [text[0:3], text[4], text[6], text[8:11], text[12], text[17:21], text[23:26]] + delimit_by_length(text[28:],5)
        else:
            return [text]

# Parse every line in the rotation
list_to_df = []
for k in rotation_dict.keys():
    for name in route_list:
        if name in k:
            route_name = name
    route_header = [k, route_name, division_dict[route_name]] + rotation_dict_headers[k]
    list_to_df.append(route_header)
    for line in rotation_dict[k]:
        data = [k, route_name, division_dict[route_name]] + parserotation(line)
        list_to_df.append(data)

#print(rotation_dict[k])
#print(parserotation(rotation_dict[k][0]))

# Calculate how many columns the dataframe needs
fixed_columns = ['Route_Name', 'Route','Division','LineNum','T','AMPM','EXC','Run_Start','TRAN_NUM','RUN_NUM']
sample_columns = fixed_columns.copy()

total_n = 0
for x in list_to_df:
    if len(x) > total_n:
        total_n = len(x)
n_columns = total_n - len(fixed_columns)

for n in range(n_columns):
    sample_columns.append('timepoint' + str(n))

rotation_df = pd.DataFrame(list_to_df, columns = sample_columns)

# Create a non-time corrected CSV file
#output_file = output_name +'beforetimecorrection' + '.csv'
#rotation_df.to_csv(output_file)

def formattime(time, ap):
    return '{H}:{M}{P}'.format(H = time[:2], M = time[2:], P = ap + 'M' )

def returnopposite(ampm):
    if ampm == 'A':
        return 'P'
    elif ampm == 'P':
        return 'A'

def isxbehindy(lst, x, y):
     for index,value in enumerate(lst):
        if value == x:
            if y in lst[index:]:
                return True
            else:
                return False
        else:
            return False

def converttoAMPM(row, ap_marker): # Convert each item in a list to a time
    if ap_marker == 'X':
        ap_marker = 'A'
    # Strip out blanks
    row = [r if r != '     ' else None for r in row]
    # Check to see if 11AM or 11PM is in the row
    hours = [t[:2] for t in row if t != None]
    if isxbehindy(hours,'11','12') == False:
        converted = [formattime(x,ap_marker) if x != None and x != '**' else x for x in row]
    else:
        # If time goes from 11 to 12, switch AM to PM and vice versa
        converted = []
        am_or_pm = ap_marker
        for v in row:
            if v != None and v != '**':
                if v[:2] == '12':
                    am_or_pm = returnopposite(ap_marker)
                    converted.append(formattime(v,am_or_pm))
                else:
                    converted.append(formattime(v,am_or_pm))
            else:
                converted.append(v)
        if len(converted) != len(row):
            print('ERROR', row)
    return [c.strip() if c != None else c for c in converted]

    ''' DEBUG SCRIPT
sample = rotation_df.loc[4833:4834]

converted_times = {}
schedule_start = len(fixed_columns) + 1
for row in sample.itertuples():
    if row.LineNum == 'LN NO':
        converted_times[row[0]] = row[schedule_start:]
    else:
        converted_times[row[0]] = converttoAMPM(row[schedule_start:], row.AMPM)

col = sample_columns[len(fixed_columns):]
updated_time_df = pd.DataFrame.from_dict(converted_times, orient = 'index', columns = col)

sampledf = sample.iloc[:,:10].join(updated_time_df)
'''

converted_times = {}
schedule_start = len(fixed_columns) + 1
for row in rotation_df.itertuples():
    if row.LineNum == 'LN NO':
        converted_times[row[0]] = row[schedule_start:]
    else:
        converted_times[row[0]] = converttoAMPM(row[schedule_start:], row.AMPM)

col = sample_columns[len(fixed_columns):]
updated_time_df = pd.DataFrame.from_dict(converted_times, orient = 'index', columns = col)

rotation_df_final = rotation_df.iloc[:,:10].join(updated_time_df)
rotation_df_final.head()

output_file = output_name + '.csv'
rotation_df_final.to_csv(output_file)

'''DEBUG SCRIPT
# Find broken 11s
def find11flags(row):
     changed = [str(x).strip() for x in row]
     h  = [t[:2] if t != None else t for t in changed]
     if '11' in h[11:]:
        if '*' in h or '**' in h:
            return True
     else:
        return False
    
converted_times = {}
schedule_start = len(fixed_columns) + 1
for row in rotation_df.itertuples():
    if row.LineNum == 'LN NO':
        converted_times[row[0]] = row[1:]
    else:
        if find11flags(row) == True:
            converted_times[row[0]] = row[1:]

col = sample_columns
updated_time_df = pd.DataFrame.from_dict(converted_times, orient = 'index', columns = col)
output_file = 'checkingfor11.csv'
updated_time_df.to_csv(output_file)
'''