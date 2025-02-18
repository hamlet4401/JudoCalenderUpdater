import json

import pandas as pd
import re
from datetime import datetime, date

from googleCalendar import Google

import spondApi
import asyncio
import os
import shutil


def read_excel_file(file_path):
    try:
        return pd.read_excel(file_path, header=1)
    except FileNotFoundError:
        print("File not found. Please check the file path.")
        return None
    except Exception as e:
        print("An error occurred:", e)
        return None


def extract_all_training_dates(excel_dataframe):
    today = date.today()
    year = today.year
    training_dates = []
    date_index = find_header_index(excel_dataframe, "datum")

    date_mod = 0
    date_iter = 0
    while date_iter < len(excel_dataframe.index):
        row_value = excel_dataframe.iloc[date_iter, date_index]
        if isinstance(row_value, str):
            match row_value.replace(' ', ''):
                case "MA":
                    date_mod = 3
                case "WOE":
                    date_mod = 4
                case "VRIJ":
                    date_mod = 3
                case _:
                    training_dates.append({})
            date_iter += 1
        elif isinstance(row_value, datetime):
            for x in range(date_mod):
                current_date = row_value.replace(year=year)
                training_dates.append({"start": current_date, "stop": current_date})
            date_iter += date_mod - 1
    return training_dates


def find_header_index(excel_dataframe, name):
    if excel_dataframe is not None:
        columns_lower = [col.lower() for col in excel_dataframe.columns]
        name_lower = name.lower()
        if name_lower in columns_lower:
            return columns_lower.index(name_lower)
    return None


def training_indices(excel_dataframe, name, total_trainings):
    row_indices = []
    name_index = find_header_index(excel_dataframe, name)
    for i in range(total_trainings):
        row_value = excel_dataframe.iloc[i, name_index]
        if isinstance(row_value, str) and row_value.lower() == name.lower():
            row_indices.append(i)
    return row_indices


def filter_training_dates(training_times, row_indices):
    filtered_training_times = []
    for index in row_indices:
        filtered_training_times.append(training_times[index])
    return filtered_training_times


def extract_training_hours(excel_dataframe, row_indices, total_trainings):
    filtered_training_hours = []
    hour_index = find_header_index(excel_dataframe, "uren")
    for i in range(total_trainings):
        row_value = excel_dataframe.iloc[i, hour_index]
        if i in row_indices and isinstance(row_value, str):
            filtered_training_hours.append(row_value)
    return filtered_training_hours


def update_training_times_by_time(training_times, training_hours):
    minute_zone = [0,15,30,45]
    i = 0
    while i < len(training_hours):
        time_range = re.findall(r'\d+', training_hours[i])

        prev_training = training_times[i - 1]["stop"]
        if i > 0 and prev_training.day == training_times[i]["start"].day and prev_training.hour == int(time_range[0]):
            if len(time_range) == 2:
                training_times[i - 1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
            elif len(time_range) == 3 and int(time_range[1]) not in minute_zone:
                training_times[i - 1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
            elif len(time_range) == 3 and int(time_range[1]) in minute_zone:
                training_times[i - 1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]))
            else:
                training_times[i - 1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]),
                                                                                  minute=int(time_range[3]))
            training_times.pop(i)
            training_hours.pop(i)
            i -= 1
        else:
            if len(time_range) == 2:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
            elif len(time_range) == 3 and int(time_range[1]) in minute_zone:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]))
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[1]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]))
            elif len(time_range) == 3 and int(time_range[1]) not in minute_zone:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]))
            else:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]),
                                                                                minute=int(time_range[1]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]),
                                                                              minute=int(time_range[3]))
        i += 1


def create_tmp_folder():
    """Create a tmp folder in the current directory if it doesn't exist."""
    tmp_dir = "tmp"
    if not os.path.exists(tmp_dir):
        os.makedirs(tmp_dir)
    return tmp_dir


def delete_tmp_folder():
    """Delete the tmp folder and all its contents."""
    tmp_dir = "tmp"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)


def run(name):
    google = Google()
    google.authenticate()

    tmp_dir = create_tmp_folder()

    try:
        file_path = asyncio.run(
            spondApi.get_latest_time_table(tmp_dir))

        # Process the Excel file
        excel_dataframe = read_excel_file(file_path)
        training_times = extract_all_training_dates(excel_dataframe)
        print(f"ALL TRAINING DATES")
        print(training_times)
        print()
        total_trainings = len(training_times)

        # Get the relevant row indices for the person's training data
        row_indices = training_indices(excel_dataframe, name, total_trainings)
        print(f"ROW INDICES FOR {name}'s TRAINING")
        print(row_indices)
        print()

        # Filter and extract the necessary training data
        training_times = filter_training_dates(training_times, row_indices)
        print(f"TRAINING DATES FOR {name}")
        print(training_times)
        print()
        training_hours = extract_training_hours(excel_dataframe, row_indices, total_trainings)
        update_training_times_by_time(training_times, training_hours)

        # Print the result in a pretty JSON format
        print(json.dumps(training_times, sort_keys=True, indent=4, default=str))

        # Create calendar events in Google Calendar
        for training_day in training_times:
            start_time = training_day["start"]
            stop_time = training_day["stop"]
            google.create_event(start_time, stop_time)

    finally:
        delete_tmp_folder()


if __name__ == "__main__":
    run('Hamlet')
