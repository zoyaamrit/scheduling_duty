import pandas as pd
from collections import defaultdict
import random

def read_data(file_path):
    
    excel_data = pd.read_excel(file_path, sheet_name=None)
    availability_dict = defaultdict(dict)

    for person, df in excel_data.items():
        row = 0
        while row < len(df) - 1: # check if 
            date_row = df.iloc[row]
            avail_row = df.iloc[row + 1]

            for date, availability in zip(date_row, avail_row):
                if pd.notna(date): 
                    if pd.notna(availability): # Ensure the date is not NaN
                        availability_dict[person][date] = availability
            row += 2

    # Reorganize data to be by date first 
    availability_by_date = defaultdict(lambda: {'Prefer': [], 'Available': []})

     # Reorganize the data to group by date first
    for person, dates in availability_dict.items():
        for date, status in dates.items():
            if status == 'Prefer':
                availability_by_date[date]['Prefer'].append(person)
            elif status == 'Available':
                availability_by_date[date]['Available'].append(person)

    return availability_by_date

def assign_dates(availability, max_weekdays, max_weekends):
    assigned_dates = {}
    weekend_count = defaultdict(int)
    weekday_count = defaultdict(int)
    
   
    # Assign based on availaibility 
    for date, people in availability.items():
        is_weekday = date not in ['Saturday', 'Sunday']

        # check who had the previous date 
        if assigned_dates and assigned_dates[max(assigned_dates)] != 'unassigned':
            prev = assigned_dates[max(assigned_dates)] 
        else:
            prev = None

        # Do not consider people who are unavailable
        prefer = people['Prefer'][:]
        available = people['Available'][:]

        # remove prev from consideration for this date
        # unless only 1 person is prefer/available
        if prev and (len(prefer) + len(available)) > 1:
            if prev in prefer:
                prefer.remove(prev)
            elif prev in available:
                available.remove(prev)

        # Eliminate people if they've already maxed out assigned days 
        if is_weekday:
            eligible_prefer = [p for p in prefer if weekday_count[p] <= max_weekdays]
            eligible_available = [p for p in available if weekday_count[p] <= max_weekdays]
        else:
            eligible_prefer = [p for p in prefer if weekend_count[p] <= max_weekends]
            eligible_available = [p for p in available if weekend_count[p] <= max_weekends]

        candidates = eligible_prefer + eligible_available
        random.shuffle(candidates)
        selected_person = candidates[0] if candidates else 'unassigned'
        assigned_dates[date] = selected_person  # Assign one person per day

        # Update counts for the selected person
        if selected_person != 'unassigned':
            if is_weekday:
                weekday_count[selected_person] += 1
            else:
                weekend_count[selected_person] += 1
    
    return assigned_dates


input_path = input("Please enter the input file path: ")

# Ask the user for the maximum number of weekend assignments
max_weekend = int(input("Enter the maximum number of weekEND assignments: "))

# Ask the user for the maximum number of weekday assignments
max_weekday = int(input("Enter the maximum number of weekDAY assignments: "))

availability = read_data(input_path)
schedule_arr = assign_dates(availability, max_weekday, max_weekend)

data = [{'Date': date, 'Assigned to': name} for date, name in schedule_arr.items()]
df = pd.DataFrame(data)
df_sorted = df.sort_values(by='Date')
output_file = './schedule.xlsx'  # Specify your desired output path

with pd.ExcelWriter(output_file) as writer:
    df_sorted.to_excel(writer, sheet_name='Schedule', index=False)

print(f"Schedule built with maximum {max_weekday} weekdays and {max_weekend} weekends")