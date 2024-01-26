import pandas as pd

# Step 1: Take the file as input
xlsx_file_path = './Assignment_Timecard.xlsx'

# Step 2: Programmatically analyze the file
# Read the Excel file into a pandas DataFrame
df = pd.read_excel(xlsx_file_path)

# Convert 'Time' and 'Time Out' columns to datetime format
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')


# a) Employees who have worked for 7 consecutive days
consecutive_days_filter = df.groupby('Employee Name')['Time'].diff().dt.days.eq(1).groupby(df['Employee Name']).cumsum() >= 6
consecutive_days_employees = df.loc[consecutive_days_filter, ['Employee Name', 'Position ID']].drop_duplicates()

# b) Employees with less than 10 hours between shifts but greater than 1 hour
time_between_shifts_filter = df.groupby('Employee Name')['Time'].diff().dt.total_seconds().between(3600, 36000)
between_shifts_employees = df.loc[time_between_shifts_filter, ['Employee Name', 'Position ID']].drop_duplicates()

# c) Employees who have worked for more than 14 hours in a single shift
long_shift_filter = (df['Time Out'] - df['Time']).dt.total_seconds() > 14 * 3600
long_shift_employees = df.loc[long_shift_filter, ['Employee Name', 'Position ID']].drop_duplicates()

# Print the results
print("Employees who have worked for 7 consecutive days:")
print(consecutive_days_employees.to_string(index=False) + "\n")

print("Employees with less than 10 hours between shifts but greater than 1 hour:")
print(between_shifts_employees.to_string(index=False) + "\n")

print("Employees who have worked for more than 14 hours in a single shift:")
print(long_shift_employees.to_string(index=False) + "\n")

output_file_path = 'output.txt'
with open(output_file_path, 'w') as output_file:
    output_file.write("Employees who have worked for 7 consecutive days:\n")
    output_file.write(consecutive_days_employees.to_string(index=False) + "\n\n")

    output_file.write("Employees with less than 10 hours between shifts but greater than 1 hour:\n")
    output_file.write(between_shifts_employees.to_string(index=False) + "\n\n")

    output_file.write("Employees who have worked for more than 14 hours in a single shift:\n")
    output_file.write(long_shift_employees.to_string(index=False) + "\n")