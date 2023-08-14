import calendar
from datetime import datetime
from openpyxl import Workbook
wb = Workbook()

# Select the worksheet
worksheet = wb.active

nerds = ['Eric', 'Alice', 'Bob', 'Charlie', 'David', 'Joe']

# Write the names to column A
for i, name in enumerate(nerds, start=3):
    worksheet.cell(row=i, column=1, value=name)

# Write the days to rows 1 (name) & 2 (number)
# Get the days in the month and their abbreviated names
days = calendar.month_abbr[1:]
num_days = calendar.monthrange(2023, 4)
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# Write the days and their abbreviated names to rows 1 and 2, write month to A1
for i, num_days in enumerate(calendar.monthrange(2023, 4)):
    worksheet.cell(row=1, column=i,
                   value=(calendar.day_name[datetime.strptime(('2023-04-' + str((i - 1) % 7)), '%Y-%m-%d').date().weekday()])[
                         :2])
    worksheet.cell(row=2, column=i, value=i - 1)

# Save the file
wb.save("sample.xlsx")
