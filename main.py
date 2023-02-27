import os
import requests
from requests import HTTPError
import pandas
from datetime import datetime, timedelta
import traceback

def main():
    print('==================================== WELCOME ====================================')
    print('************************* Pictorial Schedule Exporter *************************') 
    print('    ################### Software Made By James Jilhaney ###################')
    print('=================================================================================')
    print('\n')
    print('What would you like to export today? (Choose a number):\n1.Export Pictorials Tomorrow.\n2.Export Pictorial Date Range.\n3.Export Pictorials Today.\n4.Export All Pictorials in Date\n\n')
    try:
        choiceStr = input('Your choice (1-4): ')
        
        if (not choiceStr.isdigit()):
            raise ValueError('\nYour choice must be a number. Closing program...')
        
        choice = int(choiceStr)

        if (choice < 1 or choice > 4):
            raise ValueError('\nYour choice must be within the range of choices specified.')
        
        if choice == 1:
            get_schedules_tomorrow()

        elif choice == 2:
            from_date = input('Format: YYYY-MM-DD\nEnter the starting date: ')
            is_valid_format_from = validate_date_format(from_date)
            
            if (is_valid_format_from == False): raise ValueError("Invalid date format. Please follow the format YYYY-MM-DD.")

            to_date = input('Format: YYYY-MM-DD\nEnter the ending date: ')
            is_valid_format_to = validate_date_format(to_date)
            if (is_valid_format_to == False): raise ValueError("Invalid date format. Please follow the format YYYY-MM-DD.")
            
            get_schedules_in_range(from_date, to_date)
            
        elif choice == 3:
            get_schedules_today()
        elif choice ==4:
            date = input('Format: YYYY-MM-DD\nEnter a date: ')
            is_valid_date = validate_date_format(date)
            if (is_valid_date == False): raise ValueError("Invalid date format. Please follow the format YYYY-MM-DD.")
            get_schedules_in_date(date)
        print('==================================== EXPORT COMPLETE ====================================')
        print('        ************************* SAY THANK YOU MASTER *************************') 
        print('    ################### Software Made By James Jilhaney ###################')
        print('=================================================================================')
        print('\n')
    except ValueError as e:
        print(e)



def validate_date_format(date_text):
    try:
        if date_text != datetime.strptime(date_text, "%Y-%m-%d").strftime('%Y-%m-%d'):
            raise ValueError
        return True
    except ValueError:
        return False

def get_schedules_in_date(date):
    try:
        url = f"https://admin.crusaderyb.com/api/v1/pictorial/students?token=00ea1da4192a2030f9ae023de3b3143ed647bbab&date_from={date}&date_to={date}"

        result = requests.get(url)
        status = result.status_code
        data = result.json()
        if (status == 'error'):
            raise HTTPError('Something went wrong with your request. Please try again.')

        pictorial_list =  data['data']
        if (len(pictorial_list) ==0):
            print("No schedules found for date specified.")
            return
        export_schedules_to_excel(pictorial_data=pictorial_list, type='date', date=date)

    except Exception as e:
        print(e)


def get_schedules_tomorrow():
    try:
        tomorrow_date = datetime.today() + timedelta(1)
        tomorrow_date = tomorrow_date.strftime('%Y-%m-%d')
        url = f"https://admin.crusaderyb.com/api/v1/pictorial/students?token=00ea1da4192a2030f9ae023de3b3143ed647bbab&date_from={tomorrow_date}&date_to={tomorrow_date}"

        result = requests.get(url)
        status = result.status_code
        data = result.json()
        if (status == 'error'):
            raise HTTPError('Something went wrong with your request. Please try again.')

        pictorial_list =  data['data']
        if (len(pictorial_list) ==0):
            print("No schedules found for date specified.")
            return
        export_schedules_to_excel(pictorial_data=pictorial_list, type='date', date=tomorrow_date)

    except Exception as e:
        print(e)

def get_schedules_today():
    try:

        today_date = datetime.today()
        today_date = today_date.strftime('%Y-%m-%d')
        url = f"https://admin.crusaderyb.com/api/v1/pictorial/students?token=00ea1da4192a2030f9ae023de3b3143ed647bbab&date_from={today_date}&date_to={today_date}"

        result = requests.get(url)
        status = result.status_code
        data = result.json()
        if (status == 'error'):
            raise HTTPError('Something went wrong with your request. Please try again.')

        pictorial_list =  data['data']
        if (len(pictorial_list) ==0):
            print("No schedules found for date specified.")
            return
        export_schedules_to_excel(pictorial_data=pictorial_list, type='date', date=today_date)

    except Exception as e:
        print(e)

def get_schedules_in_range(fromDate, toDate):
    try:

        date_from = fromDate
        date_to = toDate
        url = f"https://admin.crusaderyb.com/api/v1/pictorial/students?token=00ea1da4192a2030f9ae023de3b3143ed647bbab&date_from={date_from}&date_to={date_to}"

        result = requests.get(url)
        status = result.status_code
        data = result.json()
        if (status == 'error'):
            raise HTTPError('Something went wrong with your request. Please try again.')

        pictorial_list =  data['data']
        if (len(pictorial_list) ==0):
            print("No schedules found for date specified.")
            return
        export_schedules_to_excel(pictorial_data=pictorial_list, type='range', date_from=date_from, date_to=date_to)

    except Exception as e:
        print(e)


def export_schedules_to_excel(pictorial_data, type, date=None, date_from=None, date_to=None):

    print('Parsing data and processing excel file.')
    print('Please wait...')
    list_of_parsed_json_data = []
    
    try:

        if (type == 'date'):
            # Runs when a single date is specified
            list_of_parsed_json_data = parse_custom_student_data_map(pictorial_data)
            parsed_df = pandas.DataFrame(list_of_parsed_json_data)
            parsed_df.to_excel(f"{date}-pictorial-schedules.xlsx",index=False)

        elif (type == 'range'):
            # Runs when a date range is specified
            list_of_parsed_json_data = parse_custom_student_data_map(pictorial_data)
            parsed_df = pandas.DataFrame(list_of_parsed_json_data)
            parsed_df.to_excel(f"{date_from}-{date_to}-pictorial-schedules.xlsx")
        
    except Exception as e:
        print(traceback.format_exc())

def parse_custom_student_data_map(list):
    # based on colleges in admin panel > settings > college management
    college_list = ['College of Agriculture', 'College of Arts and Sciences', 
                        'College of Computer Studies', 'College of Engineering',
                        'College of Nursing', 'School of Business and Management',
                        'School of Education', 'School of Medicine', 'College of law', 
                        'Graduate School']
    custom_map_list = []
    pictorial_data = list
    for pictorial_obj in pictorial_data:
        student_obj = pictorial_obj['student']
        schedule_obj = student_obj['pictorial']
        college_id = schedule_obj['college_id']
        print('\n\n')
        
        # format the datetime objects using the desired format string
        schedule = parse_date(schedule_obj['date'], schedule_obj['start_time'], schedule_obj['end_time'])

        student_obj = {
             'Year' : schedule_obj['year'],
             'ID' : student_obj['university_id'],
             'Full Name' : student_obj['full_name'],
             # -3 bc API college number starts at 3 and college list index starts at 0
             'College' : college_list[college_id-3], 
             'Schedule': schedule
         }
        custom_map_list.append(student_obj)

    return custom_map_list

def parse_date(date, start_time, end_time):
    start_datetime = datetime.strptime(date + ' ' + start_time, '%Y-%m-%d %H:%M:%S')
    end_datetime = datetime.strptime(date + ' ' + end_time, '%Y-%m-%d %H:%M:%S')

    # format the start and end times as strings with the desired format
    start_time_str = start_datetime.strftime("%B %d, %Y %A, %I:%M").title() + start_datetime.strftime(" %p").upper()
    end_time_str = end_datetime.strftime("%I:%M %p").upper()

    return start_time_str + ' - ' + end_time_str




if __name__ == '__main__':
    main()