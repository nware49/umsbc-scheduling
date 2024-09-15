import pandas as pd
import random

# Load the Excel file (replace 'your_file.xlsx' with the actual file path)
file_path = 'SBM 2024 Fall Semester Tabling Form (Responses).xlsx'
df = pd.read_excel(file_path, sheet_name='Form Responses 1')

# Extract the necessary columns (availability columns G to N)
availability_columns = df.columns[6:14]  # Adjust based on the structure

# Create a list of time slots based on row 1 of these columns
time_slots = list(df.columns[6:14])

# Initialize dictionary to track shift assignments per volunteer
volunteer_shifts = {}


# Function to assign two volunteers per shift, keeping the distribution even
def assign_volunteers_evenly(availability_df, slots, num_weeks=2):
    schedule_data = []
    for _, row in availability_df.iterrows():
        volunteer_shifts[row['Email Address']] = 0

    for week in range(1, num_weeks + 1):
        for slot in slots:
            # Get volunteers available for this time slot
            available_volunteers = availability_df[availability_df[slot] == 'Yes']

            if len(available_volunteers) >= 2:
                # Sort available volunteers by the number of shifts they've already worked
                available_volunteers_sorted = available_volunteers.copy()
                available_volunteers_sorted['Shift Count'] = available_volunteers_sorted['Email Address'].map(
                    volunteer_shifts)
                available_volunteers_sorted = available_volunteers_sorted.sort_values('Shift Count')

                min_shifts_worked = available_volunteers_sorted.head(1)['Shift Count'].values[0]

                available_min_volunteers = available_volunteers_sorted[available_volunteers_sorted['Shift Count'] == min_shifts_worked]
                # if there are more than 2 volunteers with the minimum number of shifts worked, randomize the selections
                if len(available_min_volunteers) > 2:
                    available_volunteers_sorted = available_min_volunteers.sample(frac=1)

                # Choose the two volunteers with the least number of shifts
                chosen_volunteers = available_volunteers_sorted.head(2)
                first_names = chosen_volunteers['First Name'].tolist()
                last_names = chosen_volunteers['Last Name'].tolist()
                emails = chosen_volunteers['Email Address'].tolist()

                # Assign these volunteers to this shift and update their shift count
                volunteer_shifts[emails[0]] += 1
                volunteer_shifts[emails[1]] += 1

                # Append the shift details to the schedule_data list
                schedule_data.append({
                    'Week': week,
                    'Time Slot': slot,
                    'Volunteer 1 Name': f"{first_names[0]} {last_names[0]}",
                    'Volunteer 1 Email': emails[0],
                    'Volunteer 2 Name': f"{first_names[1]} {last_names[1]}",
                    'Volunteer 2 Email': emails[1]
                })

    return schedule_data


# Assign volunteers based on the availability and even distribution
schedule_data = assign_volunteers_evenly(df, time_slots)

# Convert the schedule data to a DataFrame
schedule_df = pd.DataFrame(schedule_data)

# Split the data by week into two separate DataFrames
week1_df = schedule_df[schedule_df['Week'] == 1]
week2_df = schedule_df[schedule_df['Week'] == 2]


def save_schedule_to_excel(schedule_df, output_file):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Save the main schedule sheet
        schedule_df.to_excel(writer, sheet_name='Schedule', index=False)

        # Create a second sheet with transposed data
        transposed_data = []

        # Get unique time slots
        time_slots = schedule_df['Time Slot'].unique()

        # Create a dictionary to store volunteers by time slot
        slot_dict = {slot: [] for slot in time_slots}

        for slot in time_slots:
            slot_data = schedule_df[schedule_df['Time Slot'] == slot]
            for i, row in slot_data.iterrows():
                # Append volunteers for the current slot
                slot_dict[slot].append(row['Volunteer 1 Name'])
                slot_dict[slot].append(row['Volunteer 2 Name'])

        # Flatten the dictionary to create rows for each time slot
        max_volunteers = max(len(volunteers) for volunteers in slot_dict.values())
        transposed_data.append(['Time Slot'] + [slot for slot in time_slots])

        for i in range(max_volunteers):
            row = [f'Volunteer {i + 1}']
            for slot in time_slots:
                if i < len(slot_dict[slot]):
                    row.append(slot_dict[slot][i])
                else:
                    row.append('')
            transposed_data.append(row)

        # Convert transposed data to a DataFrame
        transposed_df = pd.DataFrame(transposed_data[1:], columns=transposed_data[0])

        # Save the transposed schedule to the second sheet
        transposed_df.to_excel(writer, sheet_name='Transposed Schedule', index=False)


# Save each week's schedule to a separate file
save_schedule_to_excel(week1_df, 'week_1_schedule.xlsx')
save_schedule_to_excel(week2_df, 'week_2_schedule.xlsx')

# Print the volunteer shift counts at the end
print("Volunteer Shift Counts:")
for volunteer, count in volunteer_shifts.items():
    print(f"{volunteer}: {count} shifts")

print("\nSchedules saved to 'week_1_schedule.xlsx' and 'week_2_schedule.xlsx'")