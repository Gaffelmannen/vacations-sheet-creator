#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import yaml
import pygsheets 
import calendar
from dateutil.parser import parse

file_credentials = "credentials/vacation-lists-66c7d9f6cb1e.json"
file_workbook="vacations-sheet"
file_sheet="2025"

def read_from_config(selected_year):
    holidays = {}
    with open("config.yml", "r") as file:
        red_days = yaml.safe_load(file)

    days = red_days["red_days"][selected_year]

    if days:
        #print(f"Holidays for {selected_year}:")
        for date, name in days.items():
            #print(f"  {selected_year}-{date}: {name}")
            holidays[f"{selected_year}-{date}"] = f"{name}"
    else:
        print(f"No holidays found for {selected_year}.")
        sys.exit(-1)

    #print(holidays)

    return holidays

def convert_to_datetime(input_str, parserinfo=None):
    return parse(input_str, parserinfo=parserinfo)

def find_first_integer(s):
    first_integer_position = -1
    for i, c in enumerate(s):
        if c.isdigit():
            first_integer_position = i
            break
    return first_integer_position
    
def add_data(red_days):
    print("Add data")

    client = pygsheets.authorize(service_account_file=file_credentials)
    spreadsht = client.open(file_workbook) 
    ws = spreadsht.worksheet("title", file_sheet) 
    
    position = 2
    year = 2025
    months = list(range(1, 13))

    print("Start")
    ws.cell("A1").set_text_format("bold", True).value = "Vacations {}".format(year)
    ws.cell("A2").value = "Weekday"
    ws.cell("A3").value = "Date"
    ws.cell("A4").value = "Note"

    for month in months:
        print(month)
        days_in_month = calendar.monthrange(year, month)[1]
        days = list(range(1, days_in_month+1))
        for day in days:
            thedate = "{0}-{1:02}-{2:02}".format(year, month, day)
            result_datetime = convert_to_datetime(thedate)
            weekday = calendar.day_name[result_datetime.weekday()]
            ws.update_value((3, position), thedate)
            ws.update_value((2, position), weekday)
            
            if weekday == "Saturday" or weekday == "Sunday":
                col_a = ws.get_col(position, returnas='cell')
                cell_number = col_a[0].label
                apply_batch_formatting(cell_number[0:find_first_integer(cell_number)])

            if thedate in red_days:
                col_a = ws.get_col(position, returnas='cell')
                cell_number = col_a[0].label
                cell_column = cell_number[0:find_first_integer(cell_number)]
                ws.cell("{}{}".format(cell_column, 4)).value = red_days[thedate]
                apply_batch_formatting(cell_column)
            
            position+=1

def format_weekend_cells():
    print("Format weekend cells")

    gc = pygsheets.authorize(service_file=file_credentials)
    spreadsht = gc.open(file_workbook)
    ws = spreadsht.worksheet("title", file_sheet) 

    cell_range = ws.range('A1:A10')
    for row in cell_range:
        for cell in row:
            cell.color = (0.8, 0.8, 0.8)

def apply_batch_formatting(column):
    print("Apply batch formatting for column {}".format(column))

    gc = pygsheets.authorize(service_file=file_credentials)
    spreadsht = gc.open(file_workbook)
    ws = spreadsht.worksheet("title", file_sheet) 
    
    requests = [
        {
            "repeatCell": {
                "range": ws.get_gridrange("{}2".format(column), "{}100".format(column)),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
    ]
    gc.sheet.batch_update(spreadsht.id, requests)
    

def apply_conditional_formatting():
    print("Apply conditional formatting")

    gc = pygsheets.authorize(service_file=file_credentials)
    spreadsht = gc.open(file_workbook)
    ws = spreadsht.worksheet("title", file_sheet) 
    ws.add_conditional_formatting('A1', 'A10', 'NUMBER_BETWEEN', {'backgroundColor':{'red':1}}, ['1','5'])
    
    print("Done")


if __name__ == "__main__":
    red_days_for_the_year = read_from_config(selected_year=file_sheet)
    add_data(red_days=red_days_for_the_year)
