import pandas as pd
import os
import csv

fc = ['Teagle Down Fitness Center', 'Helen Newman Fitness Center', 'Teagle Up Fitness Center', 'Noyes Fitness Center', 'Toni Morrison Fitness Center']
ir = ['Helen Newman Issue Room', 'Helen Newman Courts', 'Noyes Issue Room']
task_tag = {"Clos": "Closing", 
    "Switch":"Switch", 
    "Update Counts":"Counts",
    "Continuous Tasks - Facility Reset":"Facility Reset",
    "Continuous Tasks - Sanitation": "Sanitation",
    "Continuous Tasks - Member Interactions": "Member Interactions",
    "Continuous Tasks - Equipment":"Equipment",
    "Who's Here": "Switch",
    "Pre-Clos": "Laundry",
    "Before": "Opening",
    "Set Up": "Opening",
    "Opening": "Opening",
    "Heat Index": "Heat Index",
    "Court Monitor Closing": "Closing",
    "Summer Closing":"Closing",
    "Assist TU":"Location-specific",
    "Check on":"Location-specific",
    "Collect":"Location-specific",
    "Laundry":"Location-specific"
}

def assign_tags(row):
    for x in task_tag.keys():
        if row['Task Name'].startswith(x):
            return task_tag[x]
    return "Cleaning"

def calculate_completion_rate(row):
    return (1 - (row.Missed / row.Count)) * 100

def initial_transform(filename):
    input = open(filename, 'r')
    output = open("mod_"+filename, 'w')
    writer = csv.writer(output)
    reader = csv.reader(input)

    count = 0
    for r in reader:
        if(count < 13):
            count+=1
            continue
        writer.writerow(r)  
    input.close()
    output.close()
    df = pd.read_csv("mod_"+filename)
    if os.path.exists("mod_"+filename):
        os.remove("mod_"+filename)

    df = df[df['Expiration Time'].notnull()]
    df = df.drop(['User Name', 'Comments', 'Facility'],axis=1)
    df['Tag'] = df.apply(assign_tags, axis=1)

    df = df.query('Positions.str.startswith("Fitness Monitor")', engine="python")
    df = df.query('~Location.str.startswith("Appel")', engine="python")
    df = df[['Location','Positions','Task Name', 'Response', 'Tag']]
    return df

"""groupby = ['Location', x] where x = 'Task Name' or 'Tag' """
def append_completion_rate(df, groupby):
    cr_df = df.groupby(by=groupby)[[groupby[1]]].count().rename(columns={groupby[1]:"Count"})
    cr_df['Missed'] = df.groupby(by=groupby).Response.value_counts().unstack(fill_value=0).loc[:,'Missed'].tolist()
    cr_df['Completion Rate'] = cr_df.apply(calculate_completion_rate, axis=1)
    return cr_df

def append_responses(df, groupby):
    df = df.groupby(by=groupby)
    df.size().to_frame()
    df_1 = df.Response.unique().to_frame()
    df_1['Unique Responses'] = df[['Response']].nunique()['Response'].tolist()
    return df_1

def generate_task_report(filename, output):
    i_df = initial_transform(filename)

    groupby = ['Location', 'Tag']
    df_cr = append_completion_rate(i_df, groupby)
    df_r = append_responses(i_df, groupby)
    t_df = df_cr.join(df_r)

    groupby = ['Location', 'Task Name']
    df_cr = append_completion_rate(i_df, groupby)
    df_r = append_responses(i_df, groupby)
    tn_df = df_cr.join(df_r)

    with pd.ExcelWriter(output) as writer:
        t_df.to_excel(writer, sheet_name="By Tag")
        tn_df.to_excel(writer, sheet_name="By Task Name")

    print("sheet saved.")
    return True
    
# For testing purposes:
def run_locally():
    week_report = "Export (1).csv"
    generate_task_report(week_report, '')

#run_locally()