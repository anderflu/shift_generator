import pandas as pd
import numpy as np
from collections import defaultdict, Counter
import random
from datetime import datetime
from openpyxl import load_workbook

#Path til vaktliste
file_path = "Vaktønsker Lyche Kjøkken V24 uke 39 & 40.xlsx"
xls = pd.ExcelFile(file_path)

# Load the sheet into a dataframe
sheet_name = 0
df = pd.read_excel(file_path, sheet_name=sheet_name)
#print(df.head(10))


# Extract the dates into a list
date_format = '%Y-%m-%d %H:%M:%S' # Date format here must be equal to format in dataframe above
dates = df.iloc[0, 1:].dropna().tolist()
formatted_dates = [datetime.strptime(str(date), date_format).strftime('%d/%m ') for date in dates]
double_dates = [date for date in formatted_dates for _ in range(2)]
#print('Dates: ', double_dates)

#Extract the time slots
time_slots = df.iloc[1, 1::1].dropna().tolist() #Len = 16
compressed_time_slots = [time_slot[:2] + '-' + time_slot[8:10] for time_slot in time_slots]
#print('Time slots: ', compressed_time_slots)



#Create shift names
shift_names = []
for n in range(len(double_dates)):
    shift_names.append(double_dates[n] + compressed_time_slots[n])
#print(f'Shifts: {shift_names}')


# Extract relevant data from df and insert to another dataframe
start_of_chef_list = 3 
chefs_availability = df.iloc[start_of_chef_list:, :29]
chefs_availability.columns = ['Kokk'] + shift_names
#print(chefs_availability.head(10))


hangarounds_row_index = chefs_availability[chefs_availability['Kokk'].str.contains("HANGAROUNDS", case=False, na=False)].index


# Remove "AKTIVE" "HANGAROUNDS", "PANGER" AND NaN
chefs_availability = chefs_availability[
    ~chefs_availability['Kokk'].isin(["AKTIVE", "HANGAROUNDS", "PANGER"]) & 
    chefs_availability['Kokk'].notna()
]

#If Chef have not doodled => available all shifts, except from hangarounds and pangs
for index, row in chefs_availability.iterrows():
   if index >= hangarounds_row_index[0]:
       break
   if row[1:].isna().all():
       chefs_availability.loc[index, row.index[1:]] = "Kan jobbe!"

#print(chefs_availability[45:60])



week_1_df = chefs_availability.iloc[:, 0:15]
week_2_df = chefs_availability.iloc[:, [0] + list(range(15, 29))]
#print(week_1_df.head(10))
#print(week_2_df.head(10))


# Returns a lsit for available chefs on shift 'dd/mm hh-hh'
def get_available_chefs(shift, temp_df):
    return temp_df[temp_df[shift] == "Kan jobbe!"]['Kokk'].tolist()

#returns number of chefs needed for the shift
def get_num_of_chefs(shift_name):
    if '13-21' in shift_name:
        num_chefs = 3
    else:
        num_chefs = 4
    return num_chefs

""" def have_enough_old_chefs(list_of_chefs, availability_df):
    num_old_chefs = 0
    chef_column = availability_df.iloc[:, 0]
    for c in list_of_chefs:
        chef_row_index = chef_column[chef_column.str.contains(c, case=False, na=False)].index
        print(chef_row_index[0])
        val = availability_df.iloc[chef_row_index[0],1]
        if val == 1:
            num_old_chefs += 1
    
    enough_old_chefs = num_old_chefs >= 1
    
    
    return enough_old_chefs """


def assign_chefs(availability_df):

    schedule = defaultdict(list) #Dictionary for the schedule

    #Sort the shifst by availability low to high
    df_without_chefs = availability_df.drop(columns=['Kokk'])
    kan_jobbe_count = df_without_chefs.apply(lambda col: col.value_counts().get("Kan jobbe!", 0))
    kan_jobbe_count_dict = kan_jobbe_count.to_dict()

    # Sort the dictionary by the number of "Kan jobbe!" in ascending order
    sorted_kan_jobbe_count = dict(sorted(kan_jobbe_count_dict.items(), key=lambda item: item[1], reverse=False))
    sorted_shifts = list(sorted_kan_jobbe_count.keys())

    # Print the sorted dictionary
    #print(sorted_shifts)

    #Currently unused
    schedule_finished = False
    failsafe = 0

    #while not schedule_finished:
    temp_df = availability_df.copy()
    for shift in sorted_shifts:
        number_of_chefs = get_num_of_chefs(shift)
        available_chefs = get_available_chefs(shift, temp_df)
        random.shuffle(available_chefs)
        selected_chefs = available_chefs[:number_of_chefs]
        schedule[shift] = selected_chefs
        for chef in selected_chefs:
            for s in sorted_shifts:
                if s != shift:
                    temp_df.loc[temp_df['Kokk'] == chef, s] = "Opptatt"

    

    schedule_df = pd.DataFrame.from_dict(schedule, orient='index') # Load dictionary to dataframe
    schedule_df.columns = [f"Kokk {i+1}" for i in range(schedule_df.shape[1])] # Set column names
    schedule_finished = not schedule_df.isnull().values.any() # Check for empty slots in schedule

    #    failsafe += 1
    #    if failsafe == 5:
    #        break

    # Reorganize dataframe chronologically
    schedule_df.index = pd.to_datetime(schedule_df.index, format='%d/%m %H-%M')
    schedule_sorted_df = schedule_df.sort_index()
    schedule_sorted_df.index = schedule_sorted_df.index.strftime('%d/%m %H-%M')

    # Print the sorted DataFrame
    #print(schedule_sorted_df)

    return schedule_sorted_df

def check_schedule(schedule_df):
    registered_chefs = df.iloc[3:hangarounds_row_index[0], 0].dropna().tolist()
    assigned_chefs = schedule_df.values.ravel().tolist()
    assigned_chefs = [name for name in assigned_chefs if name != None]
    #print(f'Registered chefs: {len(registered_chefs)}')
    #print(f'Assigned chefs: {len(assigned_chefs)}')
    excluded_chefs = []

    # Check for duplicates
    counts = Counter(assigned_chefs)
    duplicate_chefs = [item for item, count in counts.items() if count > 1]
    if len(duplicate_chefs) == 0:
        duplicate_chefs.append('None')

    # Check for excluded chefs
    for chef in registered_chefs:
        if chef not in assigned_chefs:
            excluded_chefs.append(chef)

    if len(excluded_chefs) < 3:
        schedule_finished = True
    else:
        schedule_finished = False

    # Write to excel file
    #print(f'Excluded chefs: {excluded_chefs}')
    #print(f'Duplicated chefs: {duplicate_chefs}')

    return schedule_finished, excluded_chefs
      
week_1_finished = False
week_2_finished = False

# Re-run assign_chefs-function to minimize number of excluded chefs
""" count_1 = 0
excluded_chefs_1 = [None]*10
print("Week 1:")
while not week_1_finished:
    shift_schedule_week_1 = assign_chefs(week_1_df)
    temp_excluded_chefs_1 = excluded_chefs_1
    week_1_finished, excluded_chefs_1 = check_schedule(shift_schedule_week_1)
    if len(excluded_chefs_1) < len(temp_excluded_chefs_1):
        final_excluded_chefs_1 = excluded_chefs_1
    
    count_1 += 1
    if count_1 > 100:
        print("Could not solve restrictions")
        break

print(f"Excluded chefs: {excluded_chefs_1}")


count_2 = 0
excluded_chefs_2 = [None]*10
print("Week 2:")
while not week_2_finished:
    shift_schedule_week_2 = assign_chefs(week_2_df)
    temp_excluded_chefs_2 = excluded_chefs_2
    week_2_finished, excluded_chefs_2 = check_schedule(shift_schedule_week_2)
    if len(excluded_chefs_2) < len(temp_excluded_chefs_2):
        final_excluded_chefs_2 = excluded_chefs_2
    
    count_2 += 1
    if count_2 > 200:
        print("Could not solve restrictions")
        break

print(f"Excluded chefs: {excluded_chefs_2}") """


shift_schedule_week_1 = assign_chefs(week_1_df)
shift_schedule_week_2 = assign_chefs(week_2_df)

# Write schedules to excel files
file_path_week_1 = 'Chef_Shifts_Week_1.xlsx'
file_path_week_2 = 'Chef_Shifts_Week_2.xlsx'



# Make pretty excel-file
def save_to_file(schedule_df,file_path):
    schedule_df.to_excel(file_path, index_label = file_path[17]+'. uke')
    workbook = load_workbook(file_path)
    worksheet = workbook.active

    column_widths = [11, 30, 30, 30, 30]  # Column widths
    for i, width in enumerate(column_widths, start=1):
        worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = width

    workbook.save(file_path)

save_to_file(shift_schedule_week_1, file_path_week_1)
save_to_file(shift_schedule_week_2, file_path_week_2)

print(f"Schedule saved to {file_path_week_1} and {file_path_week_2}")


#TODO
# Håndtere ekskluderte kokker
# Implementere følgende restriksjoner:
#   Alle vakter skal være fylt opp om mulig
#   Alle aktive kokker skal være inkludert med mindre de er helt opptatt
# Gjøre output-fil finere og mer oversiktlig
# Legge inn håndtering av 'Kan om nødvendig'
# Skille mellom nye og gamle kokker
# Håndtering av vanlige errors
# Fikse slik at det fungerer med ulike antall sheets og brøkdeler av uker