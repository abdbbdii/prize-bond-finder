#          __        ______  ____  __ 
#   ____ _/ /_  ____/ / __ \/ __ \/ / 
#  / __ `/ __ \/ __  / / / / / / / /  
# / /_/ / /_/ / /_/ / /_/ / /_/ / /___
# \__,_/_.___/\__,_/\___\_\____/_____/
                                    
"""### abd Quality-of-Life functions.
This file contains different small functions for different kinds of objects that will make your life easier.
"""

from tkinter import ttk
import os, psutil, subprocess, winreg, pyautogui, requests, re


def split_text(text):
    pattern = r"[, \n\t]+"
    result = re.split(pattern, text)
    return result


def convert_to_folder_path(file_path):
    folder_path = re.sub(r"[^/\\]+$", "", file_path)
    return folder_path


def clean_csv_by_col_list(data: list):
    cleaned_rows = [[] for _ in range(len(data[0]) if data else 0)]
    # Clean the CSV data
    for row in data:
        for i, cell in enumerate(row):
            if str(cell).strip():
                cleaned_rows[i].append(cell)
    # Determine the maximum length among cleaned columns
    max_length = max(len(column) for column in cleaned_rows)
    # Pad shorter columns with empty strings to match the maximum length
    cleaned_rows = [column + [""] * (max_length - len(column)) for column in cleaned_rows]
    # Transpose
    return list(map(list, zip(*cleaned_rows)))


def clean_csv_by_col_dict(data_list):
    # Extract column names
    column_names = list(data_list[0].keys()) if data_list else []
    cleaned_columns = {col: [] for col in column_names}
    # Clean the CSV data
    for row in data_list:
        for col in column_names:
            cell = row[col]
            if str(cell).strip():
                cleaned_columns[col].append(cell)
    # Determine the maximum length among cleaned columns
    max_length = max(len(column) for column in cleaned_columns.values())
    # Pad shorter columns with empty strings to match the maximum length
    cleaned_columns = {col: column + [""] * (max_length - len(column)) for col, column in cleaned_columns.items()}
    # Transpose
    return [dict(zip(cleaned_columns.keys(), values)) for values in zip(*cleaned_columns.values())]


def combine_lists(old_data, new_values, key):
    """insert new list with key to old table"""
    updated_data = [data_dict for data_dict in old_data]
    for value in new_values:
        new_dict = {k: (value if k == key else "") for k in old_data[0]}
        updated_data.append(new_dict)
    return updated_data


def clean_csv_dict_values(data: list, length_of_values: int = 6, format_with: str = "0") -> list:
    result = []
    for line in data:
        keys = list(line.keys())
        values = list(line.values())
        for key, value in zip(keys, values):
            value = str(value)
            if 0 < len(value) <= length_of_values:
                line[key] = (format_with * (length_of_values - len(value))) + value
            elif len(value) > length_of_values:
                line[key] = value[: length_of_values - len(value)]
        result.append(line)
    return result


def dictTable_to_listTable(data: dict) -> list[list]:
    result = [list(data[0].keys())]
    for d in data:
        values = list(d.values())
        result.append(values)
    return result


def listTable_to_dictTable(data: list) -> list[dict]:
    keys = data[0]
    values = data[1:]
    result = []
    for value in values:
        temp = {}
        for i in range(len(keys)):
            temp[keys[i]] = value[i]
        result.append(temp)
    return result


def data_to_Treeview(data: list | tuple | set, Treeview_object: ttk.Treeview, header_anchor: str = "center", data_anchor: str = "w", col_width: int = None, min_col_width: int = None, write_header=None) -> None:
    """### Description
    Converts a list of list/dict of data to ttk.Treeview widget.
    Remember that:
    - If lists inside list then
        - It must contain only lists inside it.
        - It must contain a list of headings at 0 index and data onwards.
    - If dicts inside list then
        - It must contain only dicts inside it.
        - All dicts must have same keys.
    ----------
    ### Parameters
    - data : list | tuple | set
    - Treeview_object : ttk.Treeview
    """
    try:
        type = isinstance(data[0], dict)
    except IndexError:
        type = None
    if type or type == None:
        if write_header:
            keys = write_header
        else:
            keys = list(data[0].keys())
        Treeview_object.configure(columns=keys, show="headings")
        for key in keys:
            Treeview_object.heading(key, text=key, anchor=header_anchor)
            Treeview_object.column(key, anchor=data_anchor, minwidth=min_col_width, width=col_width)
        for i in list(range(len(data))):
            Treeview_object.insert("", index=i, values=list(data[i].values()))
    else:
        if write_header:
            keys = write_header
        else:
            keys = data[0]
        Treeview_object.configure(columns=keys, show="headings")
        for i in range(len(data)):
            if i == 0:
                for j, key in enumerate(data[i]):
                    Treeview_object.heading(key, text=data[i][j], anchor=header_anchor)
                    Treeview_object.column(key, anchor=data_anchor, minwidth=min_col_width, width=col_width)
                continue
            Treeview_object.insert("", index=i, values=data[i])


def Treeview_to_list(Treeview_object: ttk.Treeview) -> list:
    result = [[Treeview_object.heading(column)["text"] for column in Treeview_object["columns"]]]
    for child in Treeview_object.get_children():
        result.append(Treeview_object.item(child)["values"])
    return result


def Treeview_to_dict(Treeview_object: ttk.Treeview) -> list:
    result = []
    column_headers = [Treeview_object.heading(column)["text"] for column in Treeview_object["columns"]]
    for child in Treeview_object.get_children():
        row_values = {}
        for index, value in enumerate(Treeview_object.item(child)["values"]):
            try:
                row_values[column_headers[index]] = value
            except:
                pass
        result.append(row_values)
    return result


def get_appdata_path(postPath=""):
    appdata_path = os.getenv("APPDATA")
    return os.path.join(appdata_path, postPath)


def get_special_folders():
    userprofile_path = os.path.join(os.environ["USERPROFILE"])
    folders = ["Desktop", "Documents", "Music", "Pictures", "Videos"]
    if "OneDrive" in os.listdir(userprofile_path):
        onedrive_path = os.path.join(userprofile_path, "OneDrive")
        special_folders = {folder: os.path.join(onedrive_path, folder) for folder in folders}
    else:
        special_folders = {folder: os.path.join(userprofile_path, folder) for folder in folders}
    special_folders["Downloads"]= os.path.join(userprofile_path, "Downloads")
    return special_folders


def list_to_english(data: list) -> str:
    result = ""
    if len(data) == 1:
        return data[0]
    elif len(data) == 0:
        return None
    for string in data:
        if string != data[-1]:
            result += string + ", "
            continue
        result += "and " + string
    return result


def check_process(app_name) -> bool:
    list_of_processes = []
    for p in psutil.process_iter():
        list_of_processes.append(p.name().lower())
    return app_name in list_of_processes


# def open_file_in_default_and_check(file_path, check_for, max_wait=10):
#     subprocess.call("start " + file_path, shell=True)
#     wait=0
#     time.sleep(1)
#     while not check_process(check_for):
#         time.sleep(1)
#         wait+=1
#         if wait > max_wait:
#             return False


def open_file_with_app(file_path, app_path):
    subprocess.Popen([app_path, file_path])


def get_application_path(application_name):
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths") as key:
        sub_key = winreg.OpenKey(key, application_name)
        value, _ = winreg.QueryValueEx(sub_key, "")
        return value


def focus_window(window_title):
    try:
        window = pyautogui.getWindowsWithTitle(window_title)
        if window:
            window[0].activate()
            return True
        else:
            return False
    except Exception as e:
        print("Error:", str(e))
        return False


def get_public_ip():
    response = requests.get("https://api.ipify.org")
    if response.status_code == 200:
        ip_address = response.text
        return ip_address
    else:
        return None
