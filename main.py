import pandas as pd
import numpy as np
from collections import defaultdict
import random
from datetime import datetime

#Path til vaktliste
file_path = 'vaktliste2Uker.xlsx'
xls = pd.ExcelFile(file_path)

# Load the sheet into a dataframe
sheet_name = 0
df = pd.read_excel(file_path, sheet_name=sheet_name)
print(df.head(10))

#Extract chef names


# Extract the dates into a list
dates = df.iloc[0, 1:].dropna().tolist()
formatted_dates = [datetime.strptime(str(date), '%Y-%m-%d %H:%M:%S').strftime('%d/%m ') for date in dates]
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
#print(chefs_availability.head(10))

#If Chef have not doodled => available all shifts, except from hangarounds or pangs
for index, row in chefs_availability.iterrows():
   if row[1:].isna().all():
       chefs_availability.loc[index, row.index[1:]] = "Kan jobbe!"

#print(chefs_availability.head(10))

# Set column names of dataframe
chefs_availability.columns = ['Chef'] + shift_names
#print(chefs_availability.head(10))

week_1_df = chefs_availability.iloc[:, 0:15]
week_2_df = chefs_availability.iloc[:, [0] + list(range(15, 29))]
#print(week_1_df.head(10))
#print(week_2_df.head(10))




# Returns a lsit for available chefs on shift 'dd/mm hh-hh'
def get_available_chefs(shift, temp_df):
    return temp_df[temp_df[shift] == "Kan jobbe!"]['Chef'].tolist()

#returns number of chefs needed for the shift
def get_num_of_chefs(shift_name):
    if '13-21' in shift_name:
        num_chefs = 3
    else:
        num_chefs =4
    return num_chefs


def assign_chefs(availability_df): # Input parameters: week_x_df, (shift_names, lag funksjon til denne)

    schedule = defaultdict(list) #Dictionary for the schedule

    #Sort the shifst by availability low to high
    df_without_chefs = availability_df.drop(columns=['Chef'])
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
        chefs = get_available_chefs(shift, temp_df)
        random.shuffle(chefs)
        number_of_chefs = get_num_of_chefs(shift)
        selected_chefs = chefs[:number_of_chefs]
        schedule[shift] = selected_chefs
        for chef in selected_chefs:
            for s in sorted_shifts:
                if s != shift:
                    temp_df.loc[temp_df['Chef'] == chef, s] = "Opptatt"

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
    #print(df_sorted)

    return schedule_sorted_df

shift_schedule_week_1 = assign_chefs(week_1_df)
shift_schedule_week_2 = assign_chefs(week_2_df)

file_path_week_1 = 'Chef_Shifts_Week_1.xlsx'
file_path_week_2 = 'Chef_Shifts_Week_2.xlsx'
shift_schedule_week_1.to_excel(file_path_week_1, index_label = 'Skift uke 1')
shift_schedule_week_2.to_excel(file_path_week_2, index_label = 'Skift uke 2')

print(f"Schedule saved to {file_path_week_1} and {file_path_week_2}")


#TODO
# Last opp til GitHub
# Sjekke at alle kokker er inkludert og lage metode for å inkludre kokker som er til overs
# Inkludere Hangarounds og Panger som har doodlet
# Fjerne rader der navn = ''. aktive, hangarounds, panger
# Legge inn håndtering av 'Kan om nødvendig'
# Legge inn eventuelle error-meldinger
# Fikse slik at det fungerer med ulike antall sheets og brøkdeler av uker

# Trenger 51 kokker. Vi er 37. Hvor mange skal vi bli? N 