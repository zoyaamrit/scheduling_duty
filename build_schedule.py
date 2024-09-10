import pandas as pd
from collections import defaultdict
import random

def read_data():
    file_path = '/Users/tanviamrit/Desktop/Halligan/projs/ra_schedule/test.xlsx'
    
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

    prev = None if not assigned_dates else assigned_dates[max(assigned_dates)]
   
    # Now assign dates using the grouped data
    for date, people in availability.items():
        is_weekday = date not in ['Saturday', 'Sunday']

        prefer_people = people['Prefer']
        available_people = people['Available']

        if len(prefer_people) > 1 and prev in prefer_people:
            prefer_people.remove(prev)

        if len(available_people) > 1 and prev in available_people:
            available_people.remove(prev)

        if is_weekday:
            eligible_prefer_people = [p for p in prefer_people if weekday_count[p] < max_weekdays]
            eligible_available_people = [p for p in available_people if weekday_count[p] < max_weekdays]
        else:
            eligible_prefer_people = [p for p in prefer_people if weekend_count[p] < max_weekends]
            eligible_available_people = [p for p in available_people if weekend_count[p] < max_weekends]

        if eligible_prefer_people:
            candidates = eligible_prefer_people
        elif eligible_available_people:
            candidates = eligible_available_people
        else:
            assigned_dates[date] = 'unassigned'
            continue 

        # Assign a person to the date from candidates
        random.shuffle(candidates)
        selected_person = candidates[0]
        assigned_dates[date] = selected_person  # Assign one person per day

        prev = selected_person

        # Update counts for the selected person
        if is_weekday:
            weekday_count[selected_person] += 1
        else:
            weekend_count[selected_person] += 1
    
    

    return assigned_dates


availability = read_data()
schedule_arr = assign_dates(availability, 18, 10)

data = [{'Date': date, 'Assigned to': name} for date, name in schedule_arr.items()]
df = pd.DataFrame(data)
output_file = '/Users/tanviamrit/Desktop/Halligan/projs/ra_schedule/schedule.xlsx'  # Specify your desired output path

with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, sheet_name='Schedule', index=False)