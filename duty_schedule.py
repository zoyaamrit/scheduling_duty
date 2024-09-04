import pandas as pd
from collections import defaultdict
import random

def read_availability(file_path):
    availability_data = pd.read_excel(file_path, sheet_name=None)
    return availability_data

def process_availability(availability_data):
    availability = defaultdict(lambda: defaultdict(str))
    
    for person, sheet in availability_data.items():
        for date in sheet.columns:
            status = sheet[date].iloc[0]
            availability[date][person] = status
            
    return availability

def assign_dates(availability, max_weekdays, max_weekends):
    assigned_dates = {}  # Store one person per day
    weekend_count = defaultdict(int)
    weekday_count = defaultdict(int)

    for date, people in availability.items():
        if date.strftime('%A') in ['Saturday', 'Sunday']:
            is_weekday = False
        else:
            is_weekday = True

        prefer_people = [p for p, a in people.items() if a == 'Prefer']
        available_people = [p for p, a in people.items() if a == 'Available']

        # Pick candidates based on how many dates they've already been assigned to
        if is_weekday:
            eligible_prefer_people = [p for p in prefer_people if weekday_count[p] < max_weekdays]
            eligible_available_people = [p for p in available_people if weekday_count[p] < max_weekdays]
        else:
            eligible_prefer_people = [p for p in prefer_people if weekend_count[p] < max_weekends]
            eligible_available_people = [p for p in available_people if weekend_count[p] < max_weekends]

        # First, try to assign from eligible prefer people
        if eligible_prefer_people:
            candidates = eligible_prefer_people
        # If no eligible prefer people, assign from eligible available people
        elif eligible_available_people:
            candidates = eligible_available_people
        # If no eligible candidates at all, tag as 'unassigned'
        else:
            assigned_dates[date] = 'unassigned'
            continue  # Skip further processing for this date

        # Assign a person to the date from candidates
        random.shuffle(candidates)
        selected_person = candidates[0]
        assigned_dates[date] = selected_person  # Assign one person per day

        # Update counts for the selected person
        if is_weekday:
            weekday_count[selected_person] += 1
        else:
            weekend_count[selected_person] += 1

    return assigned_dates

def main(file_path):
    availability_data = read_availability(file_path)
    availability = process_availability(availability_data)
    assigned_dates = assign_dates(availability, 20, 10)

    # Prepare data for Excel
    data = [{'Date': date.strftime('%Y-%m-%d'), 'Assigned to': name} for date, name in assigned_dates.items()]

    # Create a DataFrame for the assignments
    df = pd.DataFrame(data)

    # Calculate counts for weekends and weekdays
    weekend_count = defaultdict(int)
    weekday_count = defaultdict(int)

    for date, name in assigned_dates.items():
        if name != 'unassigned':  # Skip unassigned entries
            if pd.to_datetime(date).strftime('%A') in ['Saturday', 'Sunday']:
                weekend_count[name] += 1
            else:
                weekday_count[name] += 1

    # Prepare data for counts
    count_data = [{'Name': name, 'Weekdays Assigned': weekday_count[name], 'Weekends Assigned': weekend_count[name]} 
                  for name in set(weekend_count) | set(weekday_count)]

    # Create a DataFrame for the counts
    count_df = pd.DataFrame(count_data)

    # Write to Excel file with both sheets
    output_file = '/Users/tanviamrit/Desktop/Halligan/projs/ra_schedule/schedule.xlsx'  # Specify your desired output path
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='Assignments', index=False)
        count_df.to_excel(writer, sheet_name='Counts', index=False)


if __name__ == "__main__":
    file_path = "/Users/tanviamrit/Desktop/Halligan/projs/ra_schedule/test.xlsx"  # Replace with the actual path to your Excel file
    main(file_path)
