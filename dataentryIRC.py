import openpyxl
import datetime

def find_column(workbook, worksheet, column_name):
    """ Returns column number for a column name given workbook name, worksheet name and column_name. """

    wb = openpyxl.load_workbook(workbook)
    ws = wb[worksheet]
    column_names = {}
    current = 0
    for column in ws.iter_cols(1, ws.max_column):
        column_names[column[0].value] = current
        current += 1
    return column_names[column_name]

def count(workbook, worksheet, column_number):
    """ Returns count of non-empty cells in a column given workbook name, worksheet name and column number.
    Assumes first row of column is the title and excludes first cell of column in the count. """

    wb = openpyxl.load_workbook(workbook)
    ws = wb[worksheet]
    count = 0
    for row_cell in ws.iter_rows(2, ws.max_row):
        if row_cell[column_number].value:
            count+=1
    return count

def entry_rate(workbook, worksheet, column_number):
    """ Returns average rate of data entry per day given workbook name, worksheet name, column number for 'date of entry'. 
    Assumes first row of column is the title and excludes first cell from calculation. """
    
    wb = openpyxl.load_workbook(workbook)
    ws = wb[worksheet]
    dates=[]
    for row_cell in ws.iter_rows(2, ws.max_row):
        date = row_cell[column_number].value
        if date:
             dates.append(date)

    dates = sorted(dates)
    delta = dates[-1]-dates[0]
    rate = len(dates)/delta.days
    return rate

def est_completion_date(workbook, cases_entered, rate, total_cases):
    cases_left = total_cases - cases_entered
    days_left = cases_left/rate
    end_date = datetime.datetime.now() + datetime.timedelta(days=days_left)
    return end_date  

worksheet_dict = {
    'BEATS.DUP REDCap Entry Project.xlsx': ['enter first - BEATS&DUP','enter second - BEATS only'],
    'Retrospective IRB data entry tracking.xlsx': ['All Phase 2'],
}

column_names = ['Initials', 'Date of Entry']

total_cases = {
    'BEATS.DUP REDCap Entry Project.xlsx': 201,
    'Retrospective IRB data entry tracking.xlsx': 541,
}
# all_columns stores [[workbook, worksheet, column name, column number]] for every column named "Initials" or "Date of Entry"
all_columns = []

# matching column names with column number and storing the following information in list 'all_columns' for each sheet: [workbook name, sheet name, column name, column number]
for workbook in worksheet_dict:
    for worksheet in worksheet_dict[workbook]:
        for column_name in column_names:
            column = find_column(workbook, worksheet, column_name)
            all_columns.append([workbook, worksheet, column_name, column])

for col in all_columns:
    # finding count of data entered for each sheet
    if col[2] == 'Initials':
        cases_entered = count(col[0], col[1], col[3])
        col[len(col):len(col)] = [cases_entered]
    # finding rate of entry for each sheet
    elif col[2] == 'Date of Entry':
        rate_of_entry = entry_rate(col[0],col[1],col[3])
        col[len(col):len(col)] = [rate_of_entry]


# workbook_data stores {workbook name:[cases entered, rate, total cases, completion date]}
workbook_data ={} 

for workbook in worksheet_dict:
    for col in all_columns:
        if col[0] == workbook:
            if col[0] not in workbook_data:
                if col[2] == 'Initials':
                    workbook_data[workbook]=[col[4],0,total_cases[workbook]]
                elif col[2] == 'Date of Entry':
                    workbook_data[workbook]=[0,col[4],total_cases[workbook]]
            else:
                if col[2] == 'Initials':
                    workbook_data[workbook][0:1] = [workbook_data[workbook][0] + col[4]]
                elif col[2] == 'Date of Entry':
                    workbook_data[workbook][1:2] = [workbook_data[workbook][1] + col[4]]
    workbook_data[workbook][1:2] = [(workbook_data[workbook][1]/len(worksheet_dict[workbook]))]
for workbook in workbook_data:
    comp_date = est_completion_date(workbook,workbook_data[workbook][0],workbook_data[workbook][1], workbook_data[workbook][2])
    workbook_data[workbook][len(workbook):len(workbook)] = [comp_date]

wb_data = openpyxl.load_workbook('IRC Data Entry Updates.xlsx')
ws = wb_data['Sheet1']
row = 2
for workbook in workbook_data:
    ws.cell(row = row, column = 1).value = workbook
    ws.cell(row = row, column = 2).value = workbook_data[workbook][0]
    ws.cell(row = row, column = 3).value = workbook_data[workbook][1]
    ws.cell(row = row, column = 4).value = workbook_data[workbook][2]
    ws.cell(row = row, column = 5).value = workbook_data[workbook][3]
    row += 1
wb_data.save('IRC Data Entry Updates.xlsx')




        


