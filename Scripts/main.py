from datetime import datetime
import uuid

import openpyxl
import pandas as pd


DATA_IN_PATH = 'Put the path to your incoming CSV here.'
EXCEL_PATH = 'Put the path to your Excel Workbook here.'


def main():
    input_data = pd.read_csv(DATA_IN_PATH)

    # Sort data by Date, then time, then Name
    input_data = input_data.sort_values(
        ['date', 'time', 'name']
    ).reset_index(drop=True)

    # Open the excel workbook and all worksheets
    wb = openpyxl.load_workbook(filename=EXCEL_PATH)
    in_tray = wb['In Tray']
    master_list = wb['Master List']
    appointments = wb['Appointments']
    long_term_records = wb['Long Term Records']

    # Reverse records in the in tray (most recent date first) and insert
    # them at the beginning of long term records
    in_tray_contents = list(in_tray.values)
    in_tray_contents = pd.DataFrame(
        in_tray_contents[1:], columns=in_tray_contents[0]
    ).sort_values(['date', 'time'], ascending=False).reset_index(drop=True)

    # Format dates
    if type(in_tray_contents['date'].iloc[0]) != str:
        in_tray_contents['date'] = in_tray_contents['date'].apply(
            lambda date: datetime.strftime(date, '%m/%d/%Y')
        )

    # Insert in tray contents into the beginning of long term records
    long_term_records.insert_rows(2, amount=len(in_tray_contents.index))

    for row in range(len(in_tray_contents.index)):
        long_term_records[f'A{row + 2}'] = in_tray_contents.iloc[row][0]
        long_term_records[f'B{row + 2}'] = in_tray_contents.iloc[row][1]
        long_term_records[f'C{row + 2}'] = in_tray_contents.iloc[row][2]

    # Clear in tray and insert new data
    in_tray.delete_rows(2, len(in_tray_contents.index))

    for row in range(len(input_data.index)):
        in_tray[f'A{row + 2}'] = input_data.iloc[row][0]
        in_tray[f'B{row + 2}'] = input_data.iloc[row][1]
        in_tray[f'C{row + 2}'] = input_data.iloc[row][2]

    # Put master list in a DataFrame
    master_list_contents = list(master_list.values)
    master_list_contents = pd.DataFrame(
        master_list_contents[1:], columns=master_list_contents[0]
    )

    # Left merge input_data and master_list_contents, sort the resulting 
    # dataframe, and set a multi-index
    data_with_grades = pd.merge(
        input_data, 
        master_list_contents, 
        how='left', 
        on='name'
    ).sort_values(
        ['date', 'time', 'grade', 'name']
    ).set_index(
        ['date', 'time', 'grade']
    )

    # Create a dataframe to hold the appointment scheduling data
    appointments_df = pd.DataFrame(
        columns=['date', 'time', 'group id', 'student1', 'student2', 'student3']
    )

    # Constant max group size
    MAX_GROUP_SIZE = 3

    # Iterate through dates, creating a dataframe for each
    for date in data_with_grades.index.get_level_values('date').unique():
        df_for_date = data_with_grades.loc[date]
        
        # Iterate through times, creating a dataframe for each
        for time in df_for_date.index.get_level_values('time').unique():
            df_for_time = df_for_date.loc[time]

            # Iterate through grades, creating a dataframe for each
            for grade in df_for_time.index.get_level_values('grade').unique():
                list_for_grade = df_for_time.loc[grade].values.tolist()

                # Create rows for each group
                students_in_grade = len(list_for_grade)
                
                # Handle if all students can fit in one group
                if students_in_grade <= MAX_GROUP_SIZE:
                    appointments_df = pd.concat(
                        [
                            appointments_df,
                            create_appointment_row(list_for_grade, date, time)
                        ]
                    )

                # Handle if all students cannot fit in one group
                else:
                    groups_of_3_or_under = [
                        list_for_grade[i:i+3] 
                        for i in range(0, len(list_for_grade), 3)
                    ]

                    for group in groups_of_3_or_under:
                        appointments_df = pd.concat(
                            [
                                appointments_df,
                                create_appointment_row(group, date, time)
                            ]
                        )

    # Reset appointments_df's index
    appointments_df = appointments_df.reset_index(drop=True)

    # Clear the appointments worksheet
    num_old_appointments_records = len(list(appointments.values)) - 1

    if num_old_appointments_records > 0:
        appointments.delete_rows(2, num_old_appointments_records)

    # Insert new data into the appointments page
    for row in range(len(appointments_df.index)):
        appointments[f'A{row + 2}'] = appointments_df.iloc[row][0]
        appointments[f'B{row + 2}'] = appointments_df.iloc[row][1]
        appointments[f'C{row + 2}'] = appointments_df.iloc[row][2]
        appointments[f'D{row + 2}'] = appointments_df.iloc[row][3]
        appointments[f'E{row + 2}'] = appointments_df.iloc[row][4]
        appointments[f'F{row + 2}'] = appointments_df.iloc[row][5]

    # Save the workbook
    wb.save(EXCEL_PATH)


def create_appointment_row(
    students_list: list, 
    date: str, 
    time: str
) -> pd.DataFrame:
    """
    Return a dataframe row based on the students and info passed.

    Args:
        students_list (list): A list of students for the appointment 
        (three max).
        date (str): A date for the appointment.
        time (str): A time for the appointment.

    Returns:
        pd.DataFrame: A dataframe row contianing student and time 
        information for an appointment.
    """
    # Un-nest the students list make sure that it contains three 
    # entries. Fill any empty slots with an empty string.
    if len(students_list) == 1:
        students_list = [students_list[0], '', '']

        # Handle if there was still a nested list
        if type(students_list[0]) == list:
            students_list[0] = students_list[0][0]

    elif len(students_list) == 2:
        students_list = [students_list[0][0], students_list[1][0], '']

    elif len(students_list) == 3:
        students_list = [
            students_list[0][0], 
            students_list[1][0], 
            students_list[2][0]
        ]

    else:
        raise ValueError('Incorrect students_list length.')

    return pd.DataFrame(
        [[date, time, str(uuid.uuid4()), *students_list]], 
        columns=['date', 'time', 'group id', 'student1', 'student2', 'student3']
    )


if __name__ == '__main__':
    main()