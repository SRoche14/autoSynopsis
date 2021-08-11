from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
from docx import Document

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
table_list = ''
row_list = ''
logic_list = []
statement_list = []


def open_conf():
    root.filename = filedialog.askopenfilename(initialdir="/Users/sroche/Documents", title="Select Configuration File",
                                               filetypes=(("docx files", "*.docx"), ("all files", "*.*")))

    # works for only 1 header tables
    document = Document(root.filename)

    table_count = 1
    global column_list
    global table_list
    global logic_list
    global statement_list
    global row_list
    for table in document.tables:
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df = pd.DataFrame(data)
        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
        try:
            if table_count == 1:
                column_list = df['Column Name'].tolist()
                table_list = df['Table Number'].tolist()
                row_list = df['Row Number'].tolist()
            elif table_count >= 2:
                logic_list_add = df['Logic'].tolist()
                logic_list.append(logic_list_add)
                statement_list_add = df['Statement'].tolist()
                statement_list.append(statement_list_add)
        except ReferenceError:
            print("not here")
        table_count += 1

    if isinstance(column_list, list) and isinstance(table_list, list) and isinstance(row_list, list):
        if isinstance(logic_list, list) and isinstance(statement_list, list):
            start.pack_forget()
            my_label = Label(root, text="Lists created, now generate synopses!", font=("Arial", 20))
            my_label.pack()
            return column_list, table_list, logic_list, statement_list
    else:
        messagebox.showerror("Error", "Error configuring file - lists were not created.")


def gen_synopses():
    output_statements = []
    if isinstance(column_list, list) and isinstance(table_list, list) and isinstance(row_list, list):
        if len(logic_list) > 0 and len(statement_list) > 0:
            root.folder = filedialog.askdirectory(initialdir="/Users/sroche/Documents", title="Select Synopses Folder")
            docx_list = os.listdir(root.folder)
            for doc in docx_list:
                file = root.folder + "/" + doc
                document = Document(file)
                logic_set = -1
                for grouping in logic_list:
                    filter_object = filter(lambda t: t != "", grouping)
                    grouping = list(filter_object)
                    add_statements_arr = []
                    logic_set += 1
                    logic_sub_set = -1
                    for logic in grouping:
                        logic_sub_set += 1
                        print(logic_sub_set)
                        current_statement = statement_list[logic_set][logic_sub_set]
                        logic = logic.split(',')
                        correct_num = len(logic)
                        num_correct = 0
                        for item in logic:
                            if item.strip() == '':
                                continue
                            else:
                                pieces = item.strip().split(' ')
                                column, operator, value = pieces[0], pieces[1], pieces[2]
                                table_num = table_list[int(column) - 1]
                                try:
                                    table = document.tables[int(table_num) - 1]
                                    data = [[cell.text for cell in row.cells] for row in table.rows]
                                    df = pd.DataFrame(data)
                                    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
                                    row_num = row_list[int(column) - 1]
                                    column_name = column_list[int(column) - 1]
                                    try:
                                        column_name = column_name.strip()
                                        data = df[column_name][int(row_num) - 1]
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
                                                end = index + 3
                                                section = current_statement[index:end]
                                                column = int(section[1])
                                                table_num = table_list[int(column) - 1]
                                                row_num = row_list[int(column) - 1]
                                                column_name = column_list[int(column) - 1]
                                                table = document.tables[int(table_num) - 1]
                                                data = [[cell.text for cell in row.cells] for row in table.rows]
                                                df = pd.DataFrame(data)
                                                df = df.rename(columns=df.iloc[0]).drop(df.index[0])\
                                                    .reset_index(drop=True)
                                                column_name = column_name.strip()
                                                data = df[column_name][int(row_num) - 1]
                                                current_statement = current_statement.replace(section, data)
                                            add_statements_arr.append(current_statement)

                                    except KeyError:
                                        print('here')
                                        continue
                                except ReferenceError:
                                    messagebox.showerror("Table Error", "Table does not exist!")

                    if len(grouping) == (logic_sub_set + 1):
                        output_statements.append(add_statements_arr)

                root.folder2 = filedialog.askdirectory(initialdir="/Users/sroche/Documents",
                                                       title="Select Synopses Folder")
                document = Document()
                document.add_heading('Finished Synopsis!', 0)
                for number_groups in output_statements:
                    text = ''
                    for statements in number_groups:
                        text += statements + " "
                    new_para = document.add_heading("New Paragraph:", 2)
                    p = document.add_paragraph(text)
                filepath = root.folder2 + "/New_synopsis_" + doc
                document.save(filepath)

            end = Label(root, text="Completed Making Synopses. Have fun with your new free time :)", font=("Arial", 16))
            end.pack()
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
