import xlsxwriter as xl
import pandas as pd
import numpy as np
from datetime import date, timedelta as td, datetime as dt
import pprint as pp
import json
from formats import *

xls = pd.ExcelFile(r'C:\Users\bperez\Iovino Enterprises, LLC\M007-NYCHA-Coney Island Sites - Documents\General\10 - PROGRESS BILLING\Billing\22 - August 2022\01 - Sign in Sheets\August 2022 Sign In Sheets.xlsx')
ss_data = pd.read_excel(xls, sheet_name='Sign In Sheets')
ss_data['Date'] = pd.to_datetime(ss_data['Date']).dt.date


with open('workerData.json') as f:
    worker_db = json.load(f)

def get_week_endings(start, end, day_of_week_end=0):
    '''gets the week ending periods for given start and end date using 
    the day of the week it ends'''
    res = end + td(end.weekday()) +td(day_of_week_end)
    we = []
    while (res >= start):
        we.append(res)
        res -= td(7)
    
    return list(reversed(we))


def lunch_break(company):
    '''using function level arrays, returns lunch break time as decimal of hour'''
    min30 = ['ATCO','Welkin', 'Vital','Jansons','Themis', 'Dstar','Guytec','SIG', 'Tristan', 'Triangle', 'Navillus']
    min15 =['FSE']
    
    if company in min30:
        return .5
    elif company in min15:
        return .25
    else:
        return 0 


def week_employee_dic(ss_df, company):
    '''return a dictionary of employee hours for this week ending period
    '''
    all_emp = {}
    lunch_time = lunch_break(company)
    for y in range(ss_df.shape[0]):
        emp = ss_df.iloc[y,2].upper()
        if emp not in all_emp:
            all_emp[emp] = {}
            all_emp[emp][ss_df.iloc[y,0]] = ss_df.iloc[y,5].hour + (ss_df.iloc[y,5].minute/60) - lunch_time
        else:
            all_emp[emp][ss_df.iloc[y,0]] = ss_df.iloc[y,5].hour + (ss_df.iloc[y,5].minute/60) - lunch_time

    return all_emp


def write_headings(ws, week_ending):
    '''write headings on worksheet using function level array'''
    cols_array = (['Week Ending','Employee Name','S3 (Y/N)',str(week_ending - td(6)),str(week_ending - td(5)),str(week_ending - td(4)),str(week_ending - td(3)),
    str(week_ending - td(2)),str(week_ending - td(1)),str(week_ending),'Total Hours','ST','PT',
    'Base Rate', 'PT Base Rate', 'Total Base Paid','Fringe Rate','PT Fringe Rate',
    'Total Fringe Paid','Total Pay'])

    for x in range(len(cols_array)):
        ws.write(0,x,cols_array[x])


def write_hours(ws, row, emp_dic, emp, week_ending):
    '''helper function to calculate hours and boxes to write in'''
    if (week_ending - td(6)) in emp_dic[emp]:
        ws.write(row,3, emp_dic[emp][week_ending-td(6)])
        ws.write(row+1,3, emp_dic[emp][week_ending-td(6)])
    if (week_ending - td(5)) in emp_dic[emp]:
        ws.write(row,4, emp_dic[emp][week_ending-td(5)])
        ws.write(row+1,4, emp_dic[emp][week_ending-td(5)])
    if (week_ending - td(4)) in emp_dic[emp]:
        ws.write(row,5, emp_dic[emp][week_ending-td(4)])
        ws.write(row+1,5, emp_dic[emp][week_ending-td(4)])
    if (week_ending - td(3)) in emp_dic[emp]:
        ws.write(row,6, emp_dic[emp][week_ending-td(3)])
        ws.write(row+1,6, emp_dic[emp][week_ending-td(3)])
    if (week_ending - td(2)) in emp_dic[emp]:
        ws.write(row,7, emp_dic[emp][week_ending-td(2)])
        ws.write(row+1,7, emp_dic[emp][week_ending-td(2)])
    if (week_ending - td(1)) in emp_dic[emp]:
        ws.write(row,8, emp_dic[emp][week_ending-td(1)])
        ws.write(row+1,8, emp_dic[emp][week_ending-td(1)])
    if week_ending in emp_dic[emp]:
        ws.write(row,9, emp_dic[emp][week_ending])
        ws.write(row+1,9, emp_dic[emp][week_ending])

    ws.write(row,10, f'=SUM(D{row+1}:J{row+1})')
    ws.write(row,11, f'=K{row+1}-SUMIF(D{row+1}:J{row+1},">8")+COUNTIF(D{row+1}:J{row+1},">8")*8')
    ws.write(row,12, f'=SUMIF(D{row+1}:J{row+1},">8")-COUNTIF(D{row+1}:J{row+1},">8")*8')
    

def write_worksheet(ws, emp_dic, week_ending, workbook):
    '''helper function to fill in worksheet values'''
    row = 1
    for emp in emp_dic:
        ws.write(row,0, str(week_ending))
        ws.write(row,1, str(emp))
        ws.write(row,2, str(worker_db[emp]['S3']))

        write_hours(ws, row, emp_dic, emp, week_ending)

        ws.write(row, 13, worker_db[emp]['base'], currency_format(workbook, 'white'))
        ws.write(row, 14, worker_db[emp]['pt_base'], currency_format(workbook, 'white'))
        ws.write(row, 15, f'=L{row+1}*N{row+1}+M{row+1}*O{row+1}', currency_format(workbook, 'white'))
        ws.write(row, 16, worker_db[emp]['fringe'], currency_format(workbook, 'white'))
        ws.write(row, 17, worker_db[emp]['pt_fringe'], currency_format(workbook, 'white'))
        ws.write(row, 18, f'=L{row+1}*Q{row+1}+M{row+1}*R{row+1}', currency_format(workbook, 'white'))
        ws.write(row, 19, f'=P{row+1}+S{row+1}', currency_format(workbook, 'white'))
        row += 2

    ws.write(row+1,18,'Total Pay')
    ws.write(row+1,19,f'=SUM(T2:T{row})', currency_format(workbook, 'white'))
    ws.write(row+2,18,'Total REP Pay')
    ws.write(row+2,19, f'=SUMIF(C2:C{row},"=Y",T2:T{row})', currency_format(workbook, 'white'))


def create_worksheet(ss_df, workbook, week_ending, company):
    '''creates a worksheet for the CPR spreadsheet. Creates headings
    and fill in sign in sheet hours and rates of one week ending for company'''
    ss_df = ss_df[ss_df['Date'] <= week_ending]
    ss_df = ss_df[ss_df['Date'] > (week_ending - td(7))]

    ws = workbook.add_worksheet(str(week_ending))
    write_headings(ws,week_ending)

    emp_dic = week_employee_dic(ss_df, company)
    write_worksheet(ws, emp_dic, week_ending, workbook)


def create_summary(workbook, week_endings):
    '''creates the summary page of a workbook using company week ending dates'''
    ws = workbook.add_worksheet('Summary')

    ws.write(0,0,'Week Ending')
    ws.write(0,1,'Total Pay')
    ws.write(0,2,'Total REP Pay')
    for index, week_ending in enumerate(week_endings, start=1):
        ws.write(index, 0, str(week_ending))
        ws.write(index, 1, f"=XLOOKUP(B$1,'{week_ending}'!$S:$S,'{week_ending}'!$T:$T,0)", currency_format(workbook, 'white'))
        ws.write(index, 2, f"=XLOOKUP(C$1,'{week_ending}'!$S:$S,'{week_ending}'!$T:$T,0)", currency_format(workbook, 'white'))

    ws.write(index+2, 0, 'Previous Total')
    ws.write(index+3, 0, 'This Period Total')
    ws.write(index+3, 1, f'=SUM(B$2:B{index+1})', currency_format(workbook, 'white'))
    ws.write(index+3, 2, f'=SUM(C$2:C{index+1})', currency_format(workbook, 'white'))
    
    ws.write(index+5, 0, 'Contract Total')
    ws.write(index+5, 1, f'=B{index+3}+B{index+4}', currency_format(workbook, 'white'))
    ws.write(index+5, 2, f'=C{index+3}+C{index+4}', currency_format(workbook, 'white'))
        

def not_here(ss_df, worker_db):
    '''return a dictionary of names with empty values of workers no in
    worker_db json'''
    not_here = {}
    for y in range(ss_df.shape[0]):
        if ss_df.iloc[y,2].upper() not in worker_db:
            not_here[ss_df.iloc[y,2].upper()] = None

    return not_here


def create_report(ss_df, company, day_of_the_week=0):
    '''
    ss_df is a pandas dataframe containing sign in sheet data for all companies
    company is a string that matches options in ss data
    start date is a datetime.time object
    end_date is a datetime.time object
    worker_db is a json dictionary of worker information that matches employees in ss_df
    
    creates an excel spreadsheet of total pay using a company's sign in sheet data.
    worker database json is used to calculate hourly pay and fringe benefits.'''

    ss_df = ss_df[ss_df['Company'] == company]
    week_endings = get_week_endings(date(2022, 7, 25), date(2022, 8, 24), day_of_the_week)

    if (not_here(ss_df, worker_db) == {}):
        wb = xl.Workbook(f'{company} CPR Spreadsheet August 2022.xlsx')
        for week_ending in week_endings:
            create_worksheet(ss_df, wb,week_ending, company)

        create_summary(wb, week_endings)

        wb.close()

    else:
        pp.pprint(not_here(ss_df, worker_db))

   
create_report(ss_data,'MLJ', 2)
