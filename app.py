from flask import Flask, render_template, url_for, redirect, send_file, request
import pandas as pd
import requests
import numpy as np
from datetime import date
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

app = Flask(__name__)

global last_execute
last_execute = 'September 30, 2021'


def get_date():
    # Find today's date to compare to emp lists
    today = date.today()
    today = today.strftime("%B %d, %Y")

    return today


def emp_by_client_id():
    headers = {
        'accept': 'application/json',
        'x-api-key': 'XreqEzU09o8GuvgCDR2yxagXBeYPEpzC8UBozYd6'
    }

    response = requests.get(
        'https://f9tornqbue.execute-api.us-west-2.amazonaws.com/prod/api/Employee/GetEmployeeListByClientId',
        headers=headers
    )

    response = response.json()
    response = pd.json_normalize(response)

    # Export list to excel file
    response.to_excel('templates/Emp_List.xlsx', index=False)

    last_execute = get_date()


def compare_lists():
    # Read new and previous file
    df1 = pd.read_excel('templates/Emp_List.xlsx')
    df2 = pd.read_excel('templates/Temp_Emp_List.xlsx')

    # drop any NaN cells
    df1 = df1.fillna(0)
    df2 = df2.fillna(0)

    # Put employeeID column in lists to compare
    df1_empIDs = df1['employeeID'].tolist()
    df2_empIDs = df2['employeeID'].tolist()

    # Iterate through lists looking for new additions
    new_empIDs = []
    for i in df1_empIDs:
        if i not in df2_empIDs:
            new_empIDs.append(i)
    if len(new_empIDs) == 0:
        print("No new additions")

    # Add new additions to list and export to excel
    df1_addition = []
    for i in new_empIDs:
        df1_addition.append(df1.loc[df1['employeeID'] == i])

    df1_addition = pd.DataFrame(np.concatenate(df1_addition))
    df1_addition = df1_addition.rename(columns={0: 'firstName', 1: 'lastName', 2: 'preferredName', 3: 'businessUnit',
                                                4: 'jobTitle', 5: 'reportsTo', 6: 'employeeID', 7: 'username',
                                                8: 'isActive', 9: 'peoHireDate', 10: 'erHireDate', 11: 'seniorityDate',
                                                12: 'pobStatus', 13: 'pobStatusChangedDate', 14: 'lastDayWorked'})
    df1_addition.to_excel('templates/New_Hires.xlsx', index=False)

    # Create duplicate list to cut out any inactive employees
    df1_active = df1

    # Iterate through list, dropping any inactive employees
    temp = 0
    for i in df1_active['isActive']:
        if i == 'FALSE' or i == False:
            df1_active.drop(temp, inplace=True)
        temp += 1

    # Grab the only needed columns and export to excel
    contact_list = df1_active[['firstName', 'lastName', 'preferredName', 'businessUnit', 'jobTitle']]
    contact_list.to_excel('templates/contact_list.xlsx', index=False)

    # Because Brian and I are lazy, set column width to match longest cell value
    set_col_width('templates/Emp_List.xlsx')
    set_col_width('templates/New_Hires.xlsx')
    set_col_width('templates/contact_list.xlsx')

    # Create zip file file with all excel files
    files_to_zip = ['templates/Emp_List.xlsx', 'templates/New_Hires.xlsx', 'templates/contact_list.xlsx']
    with zipfile.ZipFile('templates/Emp_List&New_Hires.zip', 'w') as zipF:
        for file in files_to_zip:
            zipF.write(file, compress_type=zipfile.ZIP_DEFLATED)


def set_col_width(x):
    # Load excel sheet created in first method
    wb = load_workbook(x)
    ws = wb['Sheet1']

    # Find row with most characters and set column width
    dim_holder = DimensionHolder(worksheet=ws)

    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=20)

    ws.column_dimensions = dim_holder

    wb.save(filename=x)


@app.route('/home', methods=['GET', 'POST'])
def home():
    try:
        return render_template('home.html', last_execute=last_execute)
    except Exception as e:
        return str(e)


@app.route('/')
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        if request.form['username'] != 'admin' or request.form['password'] != 'pass':
            error = 'Invalid login credentials. Please try again.'
        else:
            return render_template('home.html')
    return render_template('login.html', error=error)


@app.route('/background_emp_list')
def background_emp_list():
    emp_by_client_id()
    compare_lists()
    try:
        return 'nothing'
    except Exception as e:
        return str(e)


@app.route('/download_file')
def download_file():
    path = 'templates/Emp_List&New_Hires.zip'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        return str(e)


if __name__ == '__main__':
    app.run(debug=True, host='10.100.100.43')
