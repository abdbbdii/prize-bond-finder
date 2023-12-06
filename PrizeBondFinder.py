import csv, abdQOL
import tkinter as tk
from tkinter import ttk, messagebox, filedialog as fd
import sv_ttk
import os, time, pyautogui, sys
import requests, webbrowser
from bs4 import BeautifulSoup

# initializing defaults
appdata_folder_path = abdQOL.get_appdata_path(appdata_folder := "PrizeBondFinder")
appdata_session_path = os.path.join(appdata_folder_path, "session.csv")
appdata_settings_path = os.path.join(appdata_folder_path, "settings.csv")

# getting appdata folder path
settings = {}
default_padding = 10
default_button_width = 12
country_translation = {
    1: "100",
    2: "200",
    3: "750",
    4: "1500",
    5: "7500",
    6: "15000",
    7: "25000",
    8: "40000",
    9: "40000 (Premium)",
    10: "25000 (Premium)",
}


def main():
    # importing settings
    global settings
    settings = import_settings()
    # making appdata folder
    if appdata_folder not in os.listdir(abdQOL.get_appdata_path()):
        os.system(f"mkdir {appdata_folder_path}")

    # making window and main_frame
    window = tk.Tk()

    window_width = 900
    window_height = 800
    left = window.winfo_screenwidth() // 2 - window_width // 2
    top = window.winfo_screenheight() // 2 - window_height // 2
    window.geometry(f"{window_width}x{window_height}+{left}+{top}")
    window.iconbitmap = "prize_bond.ico"
    window.title("Prize Bond Finder")

    main_frame = ttk.Frame(window)
    main_frame.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)
    # making bool_vars_check variable
    bool_vars_check = []
    for _ in range(len(country_translation)):
        bool_vars_check.append(tk.BooleanVar(value=False))

    ### frame1
    # creating frame1
    frame1 = ttk.Frame(main_frame)
    frame11 = ttk.Frame(frame1)
    frame12 = ttk.Frame(frame1)
    # configuring frame1
    frame1.columnconfigure(0, weight=1)
    frame1.columnconfigure(1, weight=1)
    frame1.rowconfigure(0, weight=1)
    # placing frame1
    frame1.pack(fill="x")
    frame11.grid(column=0, row=0, sticky="nsew")
    frame12.grid(column=1, row=0, sticky="nsew")
    # creating for frame1
    add_edit_entries_button = ttk.Button(frame11, text="Add/Edit Values", style="Accent.TButton", width=default_button_width, command=lambda: entry_window(table))
    import_button = ttk.Button(frame12, text="Import", width=default_button_width, command=lambda: import_command(table))
    export_button = ttk.Button(frame12, text="Export", width=default_button_width, command=lambda: export_command(table))
    # placing inside frame1
    add_edit_entries_button.pack(side="left", padx=default_padding, pady=default_padding)
    import_button.pack(side="right", padx=default_padding, pady=default_padding)
    export_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    ### frame2
    # creating frame2
    frame2 = ttk.Frame(main_frame)
    # placing frame2
    frame2.pack(fill="x")
    # creating for frame2
    clear_all_button = ttk.Button(frame2, text="Clear All", width=default_button_width, command=lambda: clear_all_command(table))
    # placing inside frame2
    clear_all_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    ### frame3
    # creating frame3
    frame3 = ttk.Frame(main_frame)
    # configuring frame3
    frame3.columnconfigure(0, weight=1)
    frame3.columnconfigure(1, weight=0)
    frame3.rowconfigure(0, weight=1)
    frame3.rowconfigure(1, weight=0)
    # placing frame3
    frame3.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)
    # creating for frame3
    table = ttk.Treeview(frame3, columns=[], show="headings", padding=default_padding)
    x_scroll = ttk.Scrollbar(frame3, command=table.xview, orient="horizontal", style="TScrollbar")
    y_scroll = ttk.Scrollbar(frame3, command=table.yview, orient="vertical")
    table.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
    # placing inside frame3
    table.grid(row=0, column=0, sticky="ewns")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")
    ###

    import_command(table, appdata_session_path)

    ### frame4
    # creating frame4
    frame4 = ttk.Frame(main_frame)
    frame41 = ttk.Frame(frame4)
    frame42 = ttk.Frame(frame4)
    # configuring frame4
    frame4.columnconfigure(0, weight=1)
    frame4.columnconfigure(1, weight=1)
    frame4.rowconfigure(0, weight=1)
    # placing frame4
    frame4.pack(fill="x")
    frame41.grid(column=0, row=0, sticky="nsew")
    frame42.grid(column=1, row=0, sticky="nsew")
    # creating for frame4
    check_for = tk.Label(frame41, text="Check for")
    select_all_button = ttk.Button(frame42, text="Select All", width=default_button_width, command=lambda: select_all_command(bool_vars_check))
    deselect_all_button = ttk.Button(frame42, text="Deselect All", width=default_button_width, command=lambda: deselect_all_command(bool_vars_check))
    # placing inside frame4
    check_for.pack(side="left", padx=default_padding, pady=default_padding)
    select_all_button.pack(side="right", padx=default_padding, pady=default_padding)
    deselect_all_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    ### frame5
    # creating frame5
    frame5 = ttk.Frame(main_frame)
    # configuring frame5
    frame5.columnconfigure(list(range(5)), weight=1)
    frame5.rowconfigure(list(range(2)), weight=1)
    # placing frame5
    frame5.pack(fill="x")
    # creating for frame5
    check_buttons = [ttk.Checkbutton(frame5, text=country_translation[i + 1], style="Toggle.TButton", width=default_button_width, variable=bool_vars_check[i]) for i in range(len(country_translation))]
    # placing inside frame5
    for i in range(len(country_translation)):
        check_buttons[i].grid(row=i > 4, column=i % 5, sticky="ew", padx=default_padding, pady=default_padding)
    ###

    ttk.Separator(main_frame).pack(expand=True)

    ### frame6
    # creating frame6
    frame6 = ttk.Frame(main_frame)
    frame61 = ttk.Frame(frame6)
    frame62 = ttk.Frame(frame6)
    # configuring frame6
    frame6.columnconfigure((0, 1), weight=1)
    frame6.rowconfigure(0, weight=1)
    # placing frame6
    frame6.pack(fill="x")
    frame61.grid(column=0, row=0, sticky="nsew")
    frame62.grid(column=1, row=0, sticky="nsew")
    # creating for frame6
    switch_theme_button = ttk.Button(frame61, text="🌙", padding=4, command=lambda: switch_theme_command(switch_theme_button))
    donate_button = ttk.Button(frame61, text="🍵 Buy me a Coffee", style="Accent.TButton", command=donate_command)
    check_button = ttk.Button(frame62, text="Check", style="Accent.TButton", width=default_button_width, command=lambda: check_command(table, bool_vars_check))
    close_button = ttk.Button(frame62, text="Close", width=default_button_width, command=lambda: close_command(window, table))
    # placing inside frame6
    switch_theme_button.pack(side="left", padx=default_padding, pady=default_padding)
    donate_button.pack(side="left", padx=default_padding, pady=default_padding)
    check_button.pack(side="right", padx=default_padding, pady=default_padding)
    close_button.pack(side="right", padx=default_padding, pady=default_padding)
    window.protocol("WM_DELETE_WINDOW", lambda: close_command(window, table))
    ###

    sv_ttk.set_theme(settings["Mode"])
    window.mainloop()


def display_result(list_of_result_data):
    # making window and main_frame
    result_window = tk.Toplevel()
    result_window_width = 700
    result_window_height = 600
    left = result_window.winfo_screenwidth() // 2 - result_window_width // 2
    top = result_window.winfo_screenheight() // 2 - result_window_height // 2
    result_window.geometry(f"{result_window_width}x{result_window_height}+{left}+{top}")
    result_window.iconbitmap = "prize_bond.ico"
    result_window.title("Congrats!")
    result_window.grab_set()

    main_result_frame = ttk.Frame(result_window)
    main_result_frame.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)

    ### frame1
    # creating frame1
    frame1 = ttk.Frame(main_result_frame)
    # placing frame1
    frame1.pack(fill="x")
    # creating for frame1
    congrats_label = ttk.Label(frame1, text="Congrats! You won the following bond(s)")
    # placing inside frame1
    congrats_label.pack(side="left", padx=default_padding, pady=default_padding)
    ###

    ### frame2
    # creating frame2
    frame2 = ttk.Frame(main_result_frame)
    # configuring frame2
    frame2.columnconfigure(0, weight=1)
    frame2.columnconfigure(1, weight=0)
    frame2.rowconfigure(0, weight=1)
    frame2.rowconfigure(1, weight=0)
    # placing frame2
    frame2.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)
    # creating for frame2
    result_table = ttk.Treeview(frame2, columns=[], show="headings", padding=default_padding)
    result_x_scroll = ttk.Scrollbar(frame2, command=result_table.xview, orient="horizontal", style="TScrollbar")
    result_y_scroll = ttk.Scrollbar(frame2, command=result_table.yview, orient="vertical")
    result_table.configure(xscrollcommand=result_x_scroll.set, yscrollcommand=result_y_scroll.set)
    abdQOL.data_to_Treeview(list_of_result_data, Treeview_object=result_table, data_anchor="center", col_width=20, min_col_width=20)
    # placing inside frame2
    result_table.grid(row=0, column=0, sticky="ewns")
    result_y_scroll.grid(row=0, column=1, sticky="ns")
    result_x_scroll.grid(row=1, column=0, sticky="ew")
    ###

    ### frame3
    # creating frame3
    frame3 = ttk.Frame(main_result_frame)
    # placing frame3
    frame3.pack(fill="x")
    # creating for frame3
    export_button = ttk.Button(frame3, text="Export", width=default_button_width, style="Accent.TButton", command=lambda: export_command(result_table))
    close_button = ttk.Button(frame3, text="Close", width=default_button_width, command=result_window.destroy)
    # placing inside frame3
    export_button.pack(side="right", padx=default_padding, pady=default_padding)
    close_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    result_window.mainloop()


def load_excel(table):
    # making window and main_frame
    load_excel_box = tk.Toplevel()
    load_excel_box.attributes("-topmost", True)
    load_excel_box.overrideredirect(True)
    load_excel_box.attributes("-alpha", 0.75)

    load_excel_box_width = 300
    load_excel_box_height = 150
    left = (load_excel_box.winfo_screenwidth() - load_excel_box_width) - 100
    top = load_excel_box.winfo_screenheight() // 2 - load_excel_box_height // 2
    load_excel_box.geometry(f"{load_excel_box_width}x{load_excel_box_height}+{left}+{top}")

    main_result_frame = ttk.Frame(load_excel_box)
    main_result_frame.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)

    ### frame1
    # creating frame1
    frame1 = ttk.Frame(main_result_frame)
    # placing frame1
    frame1.pack(fill="both", expand=True)
    # creating for frame1
    data_reader_label = ttk.Label(frame1, text="Data reader\nClick on the 'Load Data' button to transfer the data.")
    # placing inside frame1
    data_reader_label.pack(side="left", padx=default_padding, pady=default_padding)
    ###

    ### frame2
    # creating frame2
    frame2 = ttk.Frame(main_result_frame)
    # placing frame2
    frame2.pack(fill="x")
    # creating for frame2
    load_button = ttk.Button(frame2, text="Load Data", width=default_button_width, style="Accent.TButton", command=lambda: load_command(table, load_excel_box))
    close_button = ttk.Button(frame2, text="Close Reader", width=default_button_width, command=load_excel_box.destroy)
    # placing inside frame2
    load_button.pack(side="right", padx=default_padding, pady=default_padding)
    close_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    load_excel_box.mainloop()


def entry_window(table):
    # making window and main_frame
    entry_window = tk.Toplevel()
    entry_window_width = 500
    entry_window_height = 500
    left = entry_window.winfo_screenwidth() // 2 - entry_window_width // 2
    top = entry_window.winfo_screenheight() // 2 - entry_window_height // 2
    entry_window.geometry(f"{entry_window_width}x{entry_window_height}+{left}+{top}")
    entry_window.iconbitmap = "prize_bond.ico"
    entry_window.title("Add/Edit Values")
    entry_window.grab_set()

    main_entry_frame = ttk.Frame(entry_window)
    main_entry_frame.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)

    ### frame1
    # creating frame1
    frame1 = ttk.Frame(main_entry_frame)
    frame11 = ttk.Frame(frame1)
    frame12 = ttk.Frame(frame1)
    # configuring frame1
    frame1.columnconfigure(0, weight=1)
    frame1.columnconfigure(1, weight=1)
    frame1.rowconfigure(0, weight=1)
    # placing frame1
    frame1.pack(fill="x")
    frame11.grid(column=0, row=0, sticky="nsw")
    frame12.grid(column=1, row=0, sticky="ne")
    # creating for frame1
    instruction_label = ttk.Label(frame11, text="Please enter one or more entries\nseparated by comma, space, tab or, new line.")
    edit_in_excel_button = ttk.Button(frame12, text="Add/Edit Values in Excel", style="Accent.TButton", command=lambda: edit_in_excel_command(table, entry_window))
    # placing inside frame1
    instruction_label.pack(side="left", padx=default_padding, pady=default_padding)
    edit_in_excel_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    ### frame2
    # creating frame2
    frame2 = ttk.Frame(main_entry_frame)
    Labelframe = ttk.Labelframe(frame2, text=None)
    # placing frame2
    frame2.pack(fill="both", expand=True)
    Labelframe.pack(fill="both", expand=True, padx=default_padding, pady=default_padding)
    # creating for frame2
    entries_text = tk.Text(Labelframe, width=0, height=0, border=0, undo=True, padx=default_padding, pady=default_padding, font="consolas")
    # placing inside frame1
    entries_text.pack(fill="both", expand=True, padx=default_padding // 2, pady=default_padding // 2)
    ###

    ### frame3
    # creating frame3
    frame3 = ttk.Frame(main_entry_frame)
    frame31 = ttk.Frame(frame3)
    frame32 = ttk.Frame(frame3)
    # configuring frame3
    frame3.columnconfigure(0, weight=1)
    frame3.columnconfigure(1, weight=1)
    frame3.rowconfigure(0, weight=1)
    # placing frame3
    frame3.pack(fill="x")
    frame31.grid(column=0, row=0, sticky="nsew")
    frame32.grid(column=1, row=0, sticky="nsew")
    # creating for frame3
    string_var_spin = tk.StringVar(value="Please select an option.")
    bond_combo = ttk.Combobox(frame31, values=list(country_translation.values()), textvariable=string_var_spin, state="readonly")
    add_entries_button = ttk.Button(frame32, text="+ Add", style="Accent.TButton", width=default_button_width, command=lambda: add_entries_command(entries_text, bond_combo, table))
    close_button = ttk.Button(frame32, text="Close", width=default_button_width, command=entry_window.destroy)
    # placing inside frame3
    bond_combo.pack(side="left", padx=default_padding, pady=default_padding)
    add_entries_button.pack(side="right", padx=default_padding, pady=default_padding)
    close_button.pack(side="right", padx=default_padding, pady=default_padding)
    ###

    entry_window.mainloop()


def import_settings():
    try:
        import_dict = {}
        with open(appdata_settings_path, "r") as file:
            reader = csv.DictReader(file)
            for row in reader:
                import_dict.update(row)
            if import_dict == {}:
                raise Exception
            return import_dict
    except:
        desktop_path = abdQOL.get_special_folders()["Desktop"]
        default_settings = {
            "Import file path": desktop_path,
            "Export file path": desktop_path,
            "Mode": "dark",
        }
        return default_settings


def export_settings():
    with open(appdata_settings_path, "w") as file:
        writer = csv.DictWriter(file, fieldnames=list(settings.keys()), lineterminator="\n")
        writer.writeheader()
        writer.writerow(settings)


def add_entries_command(entries_text, bond_combo, table):
    combo_key = bond_combo.get()
    if combo_key == "Please select an option.":
        messagebox.showinfo("Select an option", "Please select an option to add values to the table.")
        return
    box_values = abdQOL.split_text(entries_text.get("1.0", "end-1c"))
    table_data = abdQOL.Treeview_to_dict(table)
    table_data.append({key: "" for key in country_translation.values()})
    try:
        if combo_key not in table_data[0]:
            for row in table_data:
                row.update({combo_key: ""})
    except:
        for row in table_data:
            row.update({combo_key: ""})
    table_data = abdQOL.combine_lists(table_data, box_values, combo_key)
    table_data = abdQOL.clean_csv_by_col_dict(abdQOL.clean_csv_dict_values(table_data))
    clear_all_command(table, "yes")
    abdQOL.data_to_Treeview(table_data, table, "center", "center", 15, 15)


def switch_theme_command(button):
    global entries_text
    if sv_ttk.get_theme() == "dark":
        button.configure(text="🌞")
        sv_ttk.use_light_theme()
        settings["Mode"] = "light"
    elif sv_ttk.get_theme() == "light":
        button.configure(text="🌙")
        sv_ttk.use_dark_theme()
        settings["Mode"] = "dark"


def load_command(table, load_excel_box):
    if abdQOL.focus_window("Excel"):
        time.sleep(1)
        with pyautogui.hold("ctrl"):
            pyautogui.press("s")
        clear_all_command(table, answer="yes")
        import_command(table, import_file_path=appdata_session_path)
    else:
        load_excel_box.destroy()


def check_command(table, bool_vars_check):
    searchURL = "https://savings.gov.pk/latest/results.php"
    payload = {
        "country": None,
        "state": "all",
        "range_from": "",
        "range_to": "",
        "pb_number_list": None,
        "btnsearch": "Search",
    }
    table_values = abdQOL.Treeview_to_dict(table)
    empty_keys = []
    results = []
    for i in range(len(country_translation)):
        if bool_vars_check[i].get():
            payload["country"] = i + 1
            for j in range(len(table_values)):
                try:
                    value = table_values[j][country_translation[i + 1]]
                except KeyError:
                    empty_keys.append(country_translation[i + 1])
                    break
                if value != "":
                    payload["pb_number_list"] = f'{value},{payload["pb_number_list"]}'
            if payload["pb_number_list"] == "":
                empty_keys.append(country_translation[i + 1])
                continue
            payload["pb_number_list"] = payload["pb_number_list"]
            with requests.session() as s:
                try:
                    results.append(s.post(searchURL, data=payload))
                except:
                    messagebox.showerror("Error: No internet connection.", "Please make sure that you are connected to the internet connection.")
                    return
        payload["pb_number_list"] = ""
    if payload["country"] is None:
        messagebox.showinfo("Select an number", "Please select at least one bond number to continue.")
        return
    if empty_keys != []:
        messagebox.showwarning("Error: Entries not found.", f"No entries for bond(s) {abdQOL.list_to_english(empty_keys)} found.")
        if results == []:
            return

    list_of_result_data = []
    for i in range(len(results)):
        result_data = raw_to_list(results[i])
        if result_data is None:
            continue
        if i != 0:
            result_data = result_data[1:]
        if len(result_data) > 1:
            for data in result_data:
                list_of_result_data.append(data)

    if list_of_result_data == []:
        messagebox.showinfo("Prize not found", "No prize found for the given entries.")
    else:
        display_result(list_of_result_data)


def raw_to_list(raw):
    soup = BeautifulSoup(raw.text, "html.parser")
    try:
        table = soup.find_all("div", attrs={"class": "divTableRow"})
    except:
        return
    return [[entity.text for entity in row.find_all("div", attrs={"class": "divTableCell"})] for row in table]


def edit_in_excel_command(table, current_window):
    current_window.destroy()
    export_command(table, appdata_session_path, country_values=True)
    try:
        abdQOL.open_file_with_app(appdata_session_path, abdQOL.get_application_path("excel.exe"))
    except:
        messagebox.showerror("Request failed.", "Please make sure that you have Microsoft Excel installed in your PC.")
        return
    load_excel(table)


def import_command(table, import_file_path=None):
    if import_file_path is None:
        filetypes = (("CSV file (Comma Separated Values)", ".csv"), ("All files", ".*"))
        import_file_path = fd.askopenfilename(title="Import", initialdir=abdQOL.convert_to_folder_path(settings["Import file path"]), filetypes=filetypes)
        if import_file_path == "":
            return
        else:
            settings["Import file path"] = import_file_path
    try:
        with open(import_file_path, "r") as file:
            reader = csv.DictReader(file)
            import_dict = list(reader)
    except FileNotFoundError: return
    if table_data:= abdQOL.Treeview_to_dict(table):
        table_data.extend(import_dict)
    else: table_data = import_dict
    try: table_data = abdQOL.clean_csv_dict_values(abdQOL.clean_csv_by_col_dict(table_data))
    except ValueError: pass
    clear_all_command(table, "yes")
    abdQOL.data_to_Treeview(table_data, table, data_anchor="center", min_col_width=12, col_width=12, write_header=list(country_translation.values()))


def export_command(table, export_file_path=None, country_values=False):
    export_dict = abdQOL.Treeview_to_dict(table)
    if export_file_path is None:
        filetypes = (("CSV file (Comma Separated Values)", ".csv"), ("All files", ".*"))
        export_file_path = fd.asksaveasfile(title="Export", initialdir=abdQOL.convert_to_folder_path(settings["Export file path"]), defaultextension=".csv", filetypes=filetypes)
        if export_file_path is None:
            return
        else:
            export_file_path = export_file_path.name
            settings["Export file path"] = export_file_path
    column_headers = [table.heading(column)["text"] for column in table["columns"]]
    if country_values:
        column_headers = country_translation.values()
    try:
        with open(export_file_path, "w") as file:
            writer = csv.DictWriter(file, fieldnames=column_headers, lineterminator="\n")
            writer.writeheader()
            writer.writerows(export_dict)
    except PermissionError:
        pass


def select_all_command(bool_vars_check):
    for i in range(len(country_translation)):
        bool_vars_check[i].set(True)


def deselect_all_command(bool_vars_check):
    for i in range(len(country_translation)):
        bool_vars_check[i].set(False)


def clear_all_command(table:ttk.Treeview, answer=None):
    if answer is None:
        answer = messagebox.askquestion("Clear All", "All records inside the table will be lost.\nDo you want to continue?", icon="warning")
    if answer == "yes":
        for row_id in table.get_children():
            table.delete(row_id)
        keys = list(country_translation.values())
        table.configure(columns=keys, show="headings")
        for key in keys:
            table.heading(key, text=key, )


def close_command(window: tk.Tk, table):
    export_command(table, appdata_session_path)
    export_settings()
    window.quit()
    window.destroy()
    sys.exit()



def donate_command():
    webbrowser.open('https://www.buymeacoffee.com/abdbbdii')


main()

# pyinstaller --noconsole --onefile --windowed --collect-data sv_ttk --icon=prize_bond.ico PrizeBondFinder.py

# TODO:
# warn about entry box and excel containing invalid entries
# add loading sign checking out bonds
# scan with https://stackoverflow.com/questions/8017725/tkinter-in-python-is-it-supposed-to-be-slow-or-is-the-bottleneck-somewhere-els
