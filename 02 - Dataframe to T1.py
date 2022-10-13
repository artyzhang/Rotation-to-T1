'''
This notebook requires that a CSV file has been made from the Word Document. Upload the Word doc to Jupyter and open the 
"Rotation to Dataframe" notebook. Change the document name and output name and run to get a CSV. 
'''

import pandas as pd

document_name = '2022-07-09_Rubber Tire draft WEND Rotation_cleaned rotation.csv'
output_name = '2022-07-09_Rubber Tire draft WEND Rotation_Draft T1.csv'

rotation_df = pd.read_csv(document_name,index_col = 0)

timepointcolumns = [name for name in rotation_df.columns if name[:9] == 'timepoint']

routename_list = rotation_df['Route_Name'].unique()
route_list = rotation_df['Route'].unique()

# List of timepoints for each route
timepoint_dict = {}

for route in routename_list:
    df_row = rotation_df.query('Route_Name == "' + route + '" and LineNum == "LN NO"')
    timepoint_df = df_row[timepointcolumns]
    timepoints = [stop for stop in timepoint_df.iloc[0] if pd.notna(stop)]
    timepoint_dict[route] = timepoints

# Create a filter to get only pull out, pull in, and timepoint rows
searchforasterisk = []
for t in timepointcolumns:
    searchforasterisk.append(t + '== "**"')
rowquery = 'Run_Start != " " or ' + ' or '.join(searchforasterisk)

# Run filter
rotation_timepoints = rotation_df.query(rowquery)

# Filter out timepoint rows
rotation_runs = rotation_timepoints.query('TRAN_NUM != "TRAN NUM"')

def findroutes(train_number, df):
    # Returns routes associated with a train number
    routenames = df.loc[(df['TRAN_NUM'] == train_number)]['Route_Name'].unique()
    return routenames

def finddivisions(train_number, df):
    # Returns the divisions associated with a train number
    potentialdivisions = df.loc[(df['TRAN_NUM'] == train_number)]['Division'].unique()
    if len(potentialdivisions) > 1:
        print(train_number, ' is associated with more than one division:',' '.join([x for x in potentialdivisions]))
    return ', '.join([x for x in potentialdivisions])

def ziptimepointnames(tp_dict, route, row, column_names):
    # Takes timepoint labels from a dictionary and links them to the values in a dataframe
    timepointnames = tp_dict[route]
    scheduledata = row[column_names]
    time = [time for time in scheduledata.iloc[0]]
    return zip(timepointnames,time)

def find_pullout(train_number, df):
    # Filters the dataframe to only the first run of a train, then returns the first non-blank value found
    datarow = rotation_runs.loc[(rotation_runs['TRAN_NUM'] == train_number) & (rotation_runs['Run_Start'] == '*')]
    if len(datarow) == 0:
        return [findroutes(train_number,df),finddivisions(train_number,df),('No match','No match')]
    elif len(datarow) == 1:
        poroute = findroutes(train_number, datarow)[0]
        division = finddivisions(train_number, datarow)
        for item in ziptimepointnames(timepoint_dict, poroute, datarow, timepointcolumns):
            if pd.notnull(item[1]):
                return [poroute,division,item]
    else:
        return [findroutes(train_number,df),finddivisions(train_number,df),('Multiple matches','Check source data')]

def findnonblank(lst, matchvalue):
    match = next(i for i, j in enumerate(lst) if j == matchvalue) + 1
    for v in lst[match:]:
        if pd.isnull(v) or v == None:
            match += 1
        else:
            return match

def find_pullin(train_number, df):
    # Find the time preceding a ** value for a given train number
    datarows = rotation_runs.loc[rotation_runs['TRAN_NUM'] == train_number]
    pirows = []
    for row in datarows.itertuples():
        for i, j in enumerate(row):
            if j == "**":
                pirows.append((row[0],i))
    if len(pirows) == 0:
        return [findroutes(train_number,df),finddivisions(train_number,df),('No match','No match')]
    elif len(pirows) == 1:
        matchingrow = datarows.loc[[pirows[0][0]],:]
        piroute = matchingrow.loc[pirows[0][0],'Route_Name']
        division = matchingrow.loc[pirows[0][0],'Division']
        labeledrows = ziptimepointnames(timepoint_dict, piroute, matchingrow, timepointcolumns)
        searchlist = [y for y in labeledrows][::-1]
        timevalues = [y for (x, y) in searchlist]
        return [piroute,division,searchlist[findnonblank(timevalues,"**")]]
    else:
        print('Multiple matches for ', train_number, 'Check rows: ', ', '.join([str(x) for (x,y) in pirows]))
        return [findroutes(train_number,df),finddivisions(train_number,df),('Multiple matches','Check source data')]
    
'''TESTING BLOCK
print(rotation_runs.columns)
print(findroutes('0804',rotation_runs))
print(find_pullout('0000',rotation_runs))
print(find_pullin('9531',rotation_runs))
'''

# Create a dictionary of the train schedule
train_dict = {}
train_nums = rotation_runs['TRAN_NUM'].unique()

for train in train_nums:
    pullout_info = find_pullout(train,rotation_runs)
    pullin_info = find_pullin(train, rotation_runs)
    train_dict[train] = [pullout_info[0],
                         pullout_info[1],
                         pullout_info[2][0],
                         pullout_info[2][1],
                         pullin_info[0],
                         pullin_info[1],
                         pullin_info[2][0],
                         pullin_info[2][1]]

T1_columns = ['PO_Route','PO_Division','PO_Stop', 'PO_Time','PI_Route','PI_Division','PI_Stop','PI_Time']
T1 = pd.DataFrame.from_dict(train_dict, orient='index', columns = T1_columns)
T1.head()

T1.to_csv(output_name)