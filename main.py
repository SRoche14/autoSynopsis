from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
from docx import Document
from datetime import datetime
from dateutil import relativedelta
import re

root = Tk()
root.iconbitmap('kermit_icon.ico')
app_width = 1000
app_height = 1200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')

column_list = ''
row_list = ''
logic_list = []
statement_list = []
table_names = []
description = []
endLabel = Label(root, text="Completed Making Synopses. Have fun with your new free time :)", font=("Arial", 16))
my_label = Label(root, text="Lists created, now generate synopses!", font=("Arial", 20))


def open_conf():
    root.filename = filedialog.askopenfilename(initialdir="/Users/sroche/Documents/AutoSynopsis",
                                               title="Select Configuration File",
                                               filetypes=(("docx files", "*.docx"), ("all files", "*.*")))

    # works for only 1 header tables
    document = Document(root.filename)

    table_count = 1
    global endLabel
    global column_list
    global logic_list
    global statement_list
    global row_list
    global table_names
    global my_label
    global description
    column_list = ''
    row_list = ''
    logic_list = []
    statement_list = []
    table_names = []
    description = []
    my_label.pack_forget()
    endLabel.pack_forget()
    for table in document.tables:
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df = pd.DataFrame(data)
        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
        try:
            if table_count == 1:
                column_list = df['Column Number'].tolist()
                table_names = df['Table Name'].tolist()
                row_list = df['Row Number'].tolist()
                description = df['Description'].tolist()
            elif table_count >= 2:
                logic_list_add = df['Logic'].tolist()
                logic_list.append(logic_list_add)
                statement_list_add = df['Statement'].tolist()
                statement_list.append(statement_list_add)

        except ReferenceError:
            print("not here")
        table_count += 1

    if isinstance(column_list, list) and isinstance(table_names, list) and isinstance(row_list, list):
        if isinstance(logic_list, list) and isinstance(statement_list, list):
            start.pack_forget()

            my_label.pack()
            return column_list, table_names, logic_list, statement_list
    else:
        messagebox.showerror("Error", "Error configuring file - lists were not created.")


def gen_synopses():
    output_statements = []
    table_names_docx = []
    if isinstance(column_list, list) and isinstance(table_names, list) and isinstance(row_list, list):
        if len(logic_list) > 0 and len(statement_list) > 0:
            root.folder = filedialog.askdirectory(initialdir="/Users/sroche/Documents/AutoSynopsis",
                                                  title="Select Synopses Folder")
            docx_list = os.listdir(root.folder)
            for doc in docx_list:
                if doc == '.DS_Store' or '~' in doc:
                    continue
                file = root.folder + "/" + doc
                document = Document(file)
                for para in document.paragraphs:
                    if para.style.name == 'Detail - Heading Synopsis':
                        table_names_docx.append(para.text.strip())
                logic_set = -1
                for grouping in logic_list:
                    filter_object = filter(lambda t: t != "", grouping)
                    grouping = list(filter_object)
                    add_statements_arr = []
                    logic_set += 1
                    logic_sub_set = -1
                    for logic in grouping:
                        logic_sub_set += 1
                        current_statement = statement_list[logic_set][logic_sub_set]
                        logic = logic.split(',')
                        correct_num = len(logic)
                        num_correct = 0
                        for item in logic:
                            if item.strip().lower() == 'n/a':
                                count = 0
                                for i in current_statement:
                                    if i == '{':
                                        count = count + 1
                                for iteration in range(count):
                                    index = int(current_statement.find('{'))
                                    end = int(current_statement.find('}') + 1)
                                    section = current_statement[index:end]
                                    column = current_statement[(index + 1):(end - 1)]
                                    if "/" in column:
                                        sep = column.split("/")
                                        column = sep[0]
                                    table_title = table_names[int(column) - 1]
                                    index2 = table_names_docx.index(table_title)
                                    table_num = index2
                                    row_num = row_list[int(column) - 1]
                                    row_num = int(row_num) - 1
                                    column_name = column_list[int(column) - 1]
                                    desc = description[int(column) - 1]
                                    table = document.tables[table_num]
                                    data = [[cell.text for cell in row.cells] for row in table.rows]
                                    df = pd.DataFrame(data)
                                    df = df.rename(columns=df.iloc[0]).drop(df.index[0]) \
                                        .reset_index(drop=True)
                                    column_name = column_name.strip()
                                    column_name = int(column_name) - 1
                                    data = df.iloc[row_num, column_name]
                                    for rep in ["/mo", "/yr"]:
                                        data = data.replace(rep, "")
                                    if "How Old" in desc:
                                        separate = data.split()
                                        current_time = datetime.now()
                                        curr_yr = current_time.year
                                        curr_mon = current_time.month
                                        curr_day = current_time.day
                                        mon = separate[0]
                                        day = int(separate[1])
                                        yr = int(separate[2])
                                        datetime_object = datetime.strptime(mon, "%b")
                                        mon = datetime_object.month
                                        today = datetime(curr_yr, curr_mon, curr_day)
                                        birth = datetime(yr, mon, day)
                                        # Get the interval between two dates
                                        diff = relativedelta.relativedelta(today, birth)
                                        data = diff.years
                                    elif "Exact Age" in desc:
                                        life_expectancy = data
                                        birth_year = df.iloc[0, 1]
                                        birth_year = birth_year.split()
                                        year = birth_year[-1]
                                        life_expectancy = life_expectancy.split()
                                        dead_year = life_expectancy[-1]
                                        data = abs(int(dead_year) - int(year))

                                    elif "In Dollars" in desc:
                                        data = re.findall(r'\$.*', data)
                                        data = data[0]
                                    elif "Check Versus Now" in desc:
                                        find_year = data.split()
                                        find_month = find_year[0]
                                        datetime_object = datetime.strptime(find_month, "%b")
                                        find_month = datetime_object.month
                                        find_year = int(find_year[1])
                                        current_time = datetime.now()
                                        curr_yr = current_time.year
                                        curr_mon = current_time.month
                                        curr_day = current_time.day
                                        today = datetime(curr_yr, curr_mon, curr_day)
                                        retirement = datetime(find_year, find_month, 1)
                                        is_person_retired = retirement < today
                                        if is_person_retired:
                                            current_statement = current_statement.replace(section,
                                                                                          'You are currently retired.',
                                                                                          1)
                                            continue
                                        else:
                                            current_statement = current_statement.replace(section,
                                                                                          'You are not currently retired.',
                                                                                          1)
                                            continue
                                    if "/" in section:
                                        sec = section.replace("{", "").replace("}", "")
                                        new_thing = sec.split("/")
                                        data = data.replace("$", "").replace(",", "")
                                        data = int(data) / int(new_thing[1])
                                        data = "$" + str(data)

                                    current_statement = current_statement.replace(section, str(data), 1)
                                add_statements_arr.append(current_statement)
                            else:
                                pieces = item.strip().split(' ')
                                column, operator, value = pieces[0], pieces[1], pieces[2]
                                table_title = table_names[int(column) - 1]
                                index = table_names_docx.index(table_title)
                                table_num = index
                                try:
                                    table = document.tables[table_num]
                                    data = [[cell.text for cell in row.cells] for row in table.rows]
                                    df = pd.DataFrame(data)
                                    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
                                    row_num = row_list[int(column) - 1]
                                    row_num = int(row_num) - 1
                                    column_name = column_list[int(column) - 1]
                                    try:
                                        column_name = column_name.strip()
                                        column_name = int(column_name) - 1
                                        data = df.iloc[row_num, column_name]
                                        data = data.replace("$", "")
                                        if operator == ">":
                                            data = int(data)
                                            if data > int(value):
                                                num_correct += 1
                                            else:
                                                continue
                                        elif operator == "<":
                                            data = int(data)
                                            if data < int(value):
                                                num_correct += 1
                                        elif operator == "=":
                                            if value.lower() == "yes":
                                                compare = data.lower().strip()
                                                if compare == "yes":
                                                    num_correct += 1
                                            elif value.lower() == "no":
                                                compare = data.lower().strip()
                                                if compare == "no":
                                                    num_correct += 1
                                            elif int(data) == int(value):
                                                num_correct += 1
                                        if num_correct == correct_num:
                                            count = 0
                                            for i in current_statement:
                                                if i == '{':
                                                    count = count + 1
                                            for iteration in range(count):
                                                index = int(current_statement.find('{'))
                                                end = int(current_statement.find('}') + 1)
                                                section = current_statement[index:end]
                                                column = current_statement[(index + 1):(end - 1)]
                                                table_title = table_names[int(column) - 1]
                                                index2 = table_names_docx.index(table_title)
                                                table_num = index2
                                                row_num = row_list[int(column) - 1]
                                                row_num = int(row_num) - 1
                                                column_name = column_list[int(column) - 1]
                                                desc = description[int(column) - 1]
                                                table = document.tables[table_num]
                                                data = [[cell.text for cell in row.cells] for row in table.rows]
                                                df = pd.DataFrame(data)
                                                df = df.rename(columns=df.iloc[0]).drop(df.index[0]) \
                                                    .reset_index(drop=True)
                                                column_name = column_name.strip()
                                                column_name = int(column_name) - 1
                                                data = df.iloc[row_num, column_name]
                                                for rep in ["/mo", "/yr"]:
                                                    data = data.replace(rep, "")
                                                if "How Old" in desc:
                                                    separate = data.split()
                                                    current_time = datetime.now()
                                                    curr_yr = current_time.year
                                                    curr_mon = current_time.month
                                                    curr_day = current_time.day
                                                    mon = separate[0]
                                                    day = int(separate[1])
                                                    yr = int(separate[2])
                                                    datetime_object = datetime.strptime(mon, "%b")
                                                    mon = datetime_object.month
                                                    today = datetime(curr_yr, curr_mon, curr_day)
                                                    birth = datetime(yr, mon, day)
                                                    # Get the interval between two dates
                                                    diff = relativedelta.relativedelta(today, birth)
                                                    data = diff.years
                                                elif "Exact Age" in desc:
                                                    life_expectancy = data
                                                    birth_year = df.iloc[0, 1]
                                                    birth_year = birth_year.split()
                                                    year = birth_year[-1]
                                                    life_expectancy = life_expectancy.split()
                                                    dead_year = life_expectancy[-1]
                                                    data = abs(int(dead_year) - int(year))

                                                elif "In Dollars" in desc:
                                                    data = re.findall(r'\$.*', data)
                                                    data = data[0]
                                                elif "Check Versus Now" in desc:
                                                    find_year = data.split()
                                                    find_month = find_year[0]
                                                    datetime_object = datetime.strptime(find_month, "%b")
                                                    find_month = datetime_object.month
                                                    find_year = int(find_year[1])
                                                    current_time = datetime.now()
                                                    curr_yr = current_time.year
                                                    curr_mon = current_time.month
                                                    curr_day = current_time.day
                                                    today = datetime(curr_yr, curr_mon, curr_day)
                                                    retirement = datetime(find_year, find_month, 1)
                                                    is_person_retired = retirement < today
                                                    if is_person_retired:
                                                        current_statement = current_statement.replace(section,
                                                                                                      'You are currently retired.',
                                                                                                      1)
                                                        continue
                                                    else:
                                                        current_statement = current_statement.replace(section,
                                                                                                      'You are not currently retired.',
                                                                                                      1)
                                                        continue
                                                if "/" in section:
                                                    sec = section.replace("{", "").replace("}", "")
                                                    new_thing = sec.split("/")
                                                    data = data.replace("$", "").replace(",", "")
                                                    data = int(data) / int(new_thing[1])
                                                    data = "$" + str(data)

                                                current_statement = current_statement.replace(section, data, 1)
                                            add_statements_arr.append(current_statement)

                                    except KeyError:
                                        print('here')
                                        continue
                                    except IndexError:
                                        messagebox.showerror("Input Error",
                                                             "One of your inputs is incorrect - check that you put in the right rows and columns!")
                                    except ValueError:
                                        messagebox.showerror("Input Error",
                                                             "One of your inputs is incorrect - check that you put in the right rows and columns!")
                                except ReferenceError:
                                    messagebox.showerror("Table Error", "Table does not exist!")

                    if len(grouping) == (logic_sub_set + 1):
                        if len(add_statements_arr) > 0:
                            output_statements.append(add_statements_arr)

                root.folder2 = filedialog.askdirectory(initialdir="/Users/sroche/Documents",
                                                       title="Select Synopses Folder")
                document = Document()
                document.add_heading('Finished Synopsis!', 0)
                for number_groups in output_statements:
                    document.add_heading("New Paragraph:", 2)
                    for statements in number_groups:
                        text = ''
                        if "*bullet*" in statements:
                            statements = statements.replace("*bullet*", "")
                            split_up = statements.split("\n")
                            for statement in split_up:
                                document.add_paragraph(statement, style='List Bullet')
                        else:
                            text += statements + " "
                            document.add_paragraph(text)
                filepath = root.folder2 + "/New_synopsis_" + doc
                document.save(filepath)

            endLabel.pack()
        else:
            messagebox.showerror("Error", "Error configuring file - lists were not created.")
    else:
        messagebox.showerror("Error", "Error configuring file - lists were not created.")


title = Label(root, text="Welcome to AutoSynopsis!", font=("Arial", 25))
title.pack()

conf_btn = Button(root, text="Get configuration file", command=open_conf, relief=SUNKEN, padx=20, pady=10,
                  font=("Arial", 16))
conf_btn.pack()

gen_syn_btn = Button(root, text="Generate Synopsis", command=gen_synopses, relief=SUNKEN, padx=20, pady=10,
                     font=("Arial", 16))
gen_syn_btn.pack()

start = Label(root, text="Please configure this before generating Synopses.", font=("Arial", 20))
start.pack()

root.mainloop()
