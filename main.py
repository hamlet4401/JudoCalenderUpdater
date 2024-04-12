import json

import pandas as pd
import re
from datetime import datetime

from googleCalendar import Google


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
    training_dates = []
    date_index = find_header_index(excel_dataframe, "datum")

    date_mod = 0
    for i in range(len(excel_dataframe.index)):
        row_value = excel_dataframe.iloc[i, date_index]
        if isinstance(row_value, str):
            match row_value.replace(' ', ''):
                case "MA":
                    date_mod = 3
                case "WOE":
                    date_mod = 4
                case "VRIJ":
                    date_mod = 3
        elif isinstance(row_value, datetime):
            for x in range(date_mod):
                training_dates.append({"start": row_value, "stop": row_value})
    return training_dates


def find_header_index(excel_dataframe, name):
    if excel_dataframe is not None:
        columns_lower = [col.lower() for col in excel_dataframe.columns]
        print("HEADER")
        print(columns_lower)
        print()
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
            print(i)
    print("ROW INDICES")
    print(row_indices)
    print()
    return row_indices


def filter_training_dates(training_times, row_indices):
    filtered_training_times = []
    for index in row_indices:
        filtered_training_times.append(training_times[index])
    return filtered_training_times


def extract_training_hours(excel_dataframe, row_indices, training_times, total_trainings):
    filtered_training_hours = []
    hour_index = find_header_index(excel_dataframe, "uren")
    for i in range(total_trainings):
        row_value = excel_dataframe.iloc[i, hour_index]
        if i in row_indices and isinstance(row_value, str):
            filtered_training_hours.append(row_value)
    return filtered_training_hours


def update_training_times_by_time(training_times, training_hours):
    i = 0
    while i < len(training_hours):
        time_range = re.findall(r'\d+', training_hours[i])

        prev_training = training_times[i-1]["stop"]
        if i > 0 and prev_training.day == training_times[i]["start"].day and prev_training.hour == int(time_range[0]):
            if len(time_range) == 2:
                training_times[i-1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
            else:
                training_times[i-1]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]), minute=int(time_range[3]))
            training_times.pop(i)
            training_hours.pop(i)
            i -= 1
        else:
            if len(time_range) == 2:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[1]))
            else:
                training_times[i]["start"] = training_times[i]["start"].replace(hour=int(time_range[0]), minute=int(time_range[1]))
                training_times[i]["stop"] = training_times[i]["stop"].replace(hour=int(time_range[2]), minute=int(time_range[3]))
        i += 1


def run(file_path, name):
    google = Google()
    google.authenticate()

    excel_dataframe = read_excel_file(file_path)
    training_times = extract_all_training_dates(excel_dataframe)
    total_trainings = len(training_times)

    row_indices = training_indices(excel_dataframe, name, total_trainings)

    training_times = filter_training_dates(training_times, row_indices)
    training_hours = extract_training_hours(excel_dataframe, row_indices, training_times, total_trainings)
    update_training_times_by_time(training_times, training_hours)
    print(json.dumps(training_times,sort_keys=True, indent=4, default=str))

    for training_day in training_times:
        start_time = training_day["start"]
        stop_time = training_day["stop"]
        google.create_event(start_time, stop_time)


if __name__ == "__main__":
    year = 2024
    run('Lestabel.xlsx', 'Hamlet')

