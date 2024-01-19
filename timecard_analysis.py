
import pandas as pd
from datetime import timedelta

# Load the data
df = pd.read_excel('Assignment_Timecard.xlsx')

# Convert the 'Time' and 'Time Out' columns to datetime
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')

# Calculate the duration of each shift
df['Shift Duration'] = df['Time Out'] - df['Time']

# Sort the dataframe by 'Employee Name' and 'Time'
df = df.sort_values(by=['Employee Name', 'Time'])

# Reset index after sorting
df.reset_index(drop=True, inplace=True)

# Helper function to check if a list of dates contains 7 consecutive days
def has_seven_consecutive_days(dates):
    dates = sorted(set(dates))
    for i in range(len(dates) - 6):
        if (dates[i + 6] - dates[i]).days == 6:
            return True
    return False

# Group by 'Employee Name' and apply the consecutive days check
consecutive_days = df.groupby('Employee Name')['Time'].apply(lambda x: has_seven_consecutive_days(x.dt.date)).reset_index(name='7 Consecutive Days')

# Filter employees who worked 7 consecutive days
employees_7_consecutive_days = consecutive_days[consecutive_days['7 Consecutive Days']]

# Save the results to an Excel file
df[df['Employee Name'].isin(employees_7_consecutive_days['Employee Name'])].to_excel('Output/employees_7_consecutive_days.xlsx', index=False)

# Define a function to check for short breaks between shifts
def check_short_breaks(group):
    group = group.sort_values('Time')
    group['Time Between Shifts'] = group['Time'].shift(-1) - group['Time Out']
    # Check if any time between shifts is less than 10 hours but more than 1 hour
    conditions = (group['Time Between Shifts'] < timedelta(hours=10)) & (group['Time Between Shifts'] > timedelta(hours=1))
    return group[conditions]

# Apply the function to each group
short_breaks = df.groupby('Employee Name').apply(check_short_breaks).reset_index(drop=True)

# Save the results to an Excel file
df[df['Employee Name'].isin(short_breaks['Employee Name'])].to_excel('Output/employees_short_breaks.xlsx', index=False)

# Find employees who have worked more than 14 hours in a single shift
long_shifts = df[df['Shift Duration'] > timedelta(hours=14)]

# Save the results to an Excel file
df[df['Employee Name'].isin(long_shifts['Employee Name'])].to_excel('Output/employees_long_shifts.xlsx', index=False)
