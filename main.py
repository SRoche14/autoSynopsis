from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
from docx import Document
from datetime import datetime
from dateutil import relativedelta
import re
import numpy as np

root = Tk()
root.iconbitmap('kermit_icon.ico')
app_width = 1000
app_height = 1200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')


endLabel = Label(root, text="Completed Making Synopses. Have fun with your new free time :)", font=("Arial", 16))
my_label = Label(root, text="Lists created, now generate synopses!", font=("Arial", 20))


column_names = []
table_names = []
row_words = []
instance_var = []
everything = []
logic_list = []
statement_list = []


def open_conf(label):
    global column_names
    global table_names
    global row_words
    global instance_var
    global everything
    global logic_list
    global statement_list
    root.filename = filedialog.askopenfilename(initialdir="/Users/sroche/Documents/AutoSynopsis",
                                               title="Select Configuration File",
                                               filetypes=(("docx files", "*.docx"), ("all files", "*.*")))

    # works for only 1 header tables
    document = Document(root.filename)
    my_label.pack_forget()
    endLabel.pack_forget()
    table_count = 0
    for table in document.tables:
        data = [[cell.text for cell in row.cells] for row in table.rows]
        df = pd.DataFrame(data)
        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
        table_count += 1
        try:
            if table_count == 1:
                find_col_name = [col for col in df.columns if 'Column Name' in col]
                row_word = [col for col in df.columns if 'Row contain' in col]
                instance_col = [col for col in df.columns if 'Which instance' in col]
                everything_col = [col for col in df.columns if 'Get everything' in col]
                table_names = df['Table Name'].tolist()
                column_names = df[find_col_name].values.tolist()
                row_words = df[row_word].values.tolist()
                instance_var = df[instance_col].values.tolist()
                everything = df[everything_col].values.tolist()
            elif table_count >= 2:
                logic_list_add = df['Logic'].tolist()
                logic_list.append(logic_list_add)
                statement_list_add = df['Statement'].tolist()
                statement_list.append(statement_list_add)

        except ReferenceError:
            print("not here")
        table_count += 1

    label.pack()
    # return column_names, table_names, row_words, instance_var, everything, logic_list, statement_list


def develop_sentences(output1_arr, output2_arr):
    log_set = -1
    for logic_sets in logic_list:

        small_set = -1
        log_set += 1
        for small_log_set in logic_sets:
            small_set += 1

            logic_arr = small_log_set.split(',')
            log_holder = []
            current_fulfilled = 0
            for log_statement in logic_arr:
                if log_statement == '':
                    continue
                elif log_statement.find("If") != -1 or log_statement.find("if") != -1:
                    index = int(log_statement.find('{'))
                    end = int(log_statement.find('}') + 1)
                    section = log_statement[index:end]
                    section = section.replace("{", "").replace("}", "")
                    if "*" in section:
                        section = int(section.replace("*", "")) - 1

                        if section in output2_arr.keys():
                            current_fulfilled += 1
                        else:
                            break
                    else:
                        try:
                            section = int(section) - 1
                            log_holder.append(output1_arr[section])
                            current_fulfilled += 1
                        except KeyError:
                            break
                elif ">" or "<" or "=" in log_statement:
                    if "<" in log_statement:
                        split_log = log_statement.split("<")
                        arrow_symbol = "<"
                    elif ">" in log_statement:
                        split_log = log_statement.split(">")
                        arrow_symbol = ">"
                    elif "=" in log_statement:
                        split_log = log_statement.split("=")
                        arrow_symbol = "="
                    else:
                        break
                    decision = log_statement.count('{')
                    if decision == 2:
                        index = int(log_statement.find('{'))
                        end = int(log_statement.find('}') + 1)
                        section = log_statement[index:end]
                        section = section.replace("{", "").replace("}", "")
                        index2 = int(log_statement.find('{', log_statement.index('{') + 1))
                        end2 = int(log_statement.find('}', log_statement.index('}') + 1) + 1)
                        section2 = log_statement[index2:end2]
                        section2 = section2.replace("{", "").replace("}", "")

                        if "*" in section:
                            try:
                                section = section.replace("*", "")
                                output_use = output2_arr
                            except KeyError:
                                continue
                        else:
                            output_use = output1_arr
                        if "*" in section2:
                            try:
                                section2 = section2.replace("*", "")
                                output_use2 = output2_arr
                            except KeyError:
                                continue
                        else:
                            output_use2 = output1_arr
                        try:
                            key = int(section) - 1
                            val = output_use[key].replace("$", "").replace(",", "")
                            val = val.replace("/mo", "").replace("/yr", "").replace("/wk", "")

                            key2 = int(section2) - 1
                            val2 = output_use2[key2].replace("$", "").replace(",", "")
                            val2 = val2.replace("/mo", "").replace("/yr", "").replace("/wk", "")
                            try:
                                val = int(val)
                                val2 = int(val2)
                            except:
                                break
                            if arrow_symbol == ">":

                                if val > val2:
                                    current_fulfilled += 1
                                else:
                                    break
                            elif arrow_symbol == "<":
                                if val < val2:
                                    current_fulfilled += 1
                                else:
                                    break
                            elif arrow_symbol == "=":
                                if val == val2:
                                    current_fulfilled += 1
                                else:
                                    break
                        except:
                            break
                    elif decision == 1:
                        index = int(log_statement.find('{'))
                        end = int(log_statement.find('}') + 1)
                        section = log_statement[index:end]
                        full_section = log_statement[index:end]
                        section = section.replace("{", "").replace("}", "")
                        if "*" in section:
                            try:
                                section = section.replace("*", "")
                                output_use = output2_arr
                            except KeyError:
                                continue
                        else:
                            output_use = output1_arr
                        try:
                            key = int(section) - 1
                            val = output_use[key].replace("$", "").replace(",", "")
                            val = val.replace("/mo", "").replace("/yr", "").replace("/wk", "")
                            if val.upper() == "YES":
                                val = "yes"
                            elif val.upper() == "NO":
                                val = 'no'
                            try:
                                if val != 'yes' and val != 'no':
                                    val = int(val)
                            except:
                                break
                            if arrow_symbol == "<":
                                if full_section in split_log[0]:
                                    if val < int(split_log[1].strip()):
                                        current_fulfilled += 1
                                    else:
                                        break
                                elif full_section in split_log[1]:
                                    if val > int(split_log[0].strip()):
                                        current_fulfilled += 1
                                    else:
                                        break
                            elif arrow_symbol == ">":
                                if full_section in split_log[0]:
                                    if val > int(split_log[1].strip()):
                                        current_fulfilled += 1
                                    else:
                                        break
                                elif full_section in split_log[1]:
                                    if val < int(split_log[0].strip()):
                                        current_fulfilled += 1
                                    else:
                                        break
                            elif arrow_symbol == "=":

                                if full_section in split_log[0]:
                                    if val == 'yes':
                                        if split_log[1].strip().lower() == 'yes':
                                            current_fulfilled += 1
                                        else:
                                            break
                                    elif val == 'no':
                                        if split_log[1].strip().lower() == 'no':
                                            current_fulfilled += 1
                                        else:
                                            break
                                    else:
                                        if val == int(split_log[1].strip()):
                                            current_fulfilled += 1
                                        else:
                                            break
                                elif full_section in split_log[1]:
                                    if val == 'yes':
                                        if split_log[0].strip().lower() == 'yes':
                                            current_fulfilled += 1
                                        else:
                                            break
                                    elif val == 'no':
                                        if split_log[0].strip().lower() == 'no':
                                            current_fulfilled += 1
                                        else:
                                            break
                                    else:
                                        if val == int(split_log[0].strip()):
                                            current_fulfilled += 1
                                        else:
                                            break
                        except KeyError:
                            continue
                if current_fulfilled == len(logic_arr):
                    count = 0
                    statement = statement_list[log_set][small_set]
                    while '{' in statement and count <= 20:
                        count += 1
                        index = int(statement.find('{'))
                        end = int(statement.find('}') + 1)
                        section = statement[index:end]
                        key = section.replace('{', "").replace('}', "")
                        if "*" in key:
                            key = int(key.replace("*", "")) - 1
                            value = output2_arr[key]
                            statement = statement.replace(section, value)
                        else:
                            key = int(key) - 1
                            value = output1_arr[key]
                            statement = statement.replace(section, value)
                        if '{' and '}' not in statement:
                            print(statement)
                    if '{' not in statement and count == 0:
                        print(statement)


def gen_synopses():

    if isinstance(column_names, list) and isinstance(table_names, list) and isinstance(row_words, list):
        if len(logic_list) >= 0 and len(statement_list) >= 0:
            root.folder = filedialog.askdirectory(initialdir="/Users/sroche/Documents/AutoSynopsis",
                                                  title="Select Synopses Folder")
            docx_list = os.listdir(root.folder)
            for doc in docx_list:
                output1_list = {}
                output2_list = {}
                table_name_use = []
                for i in table_names:
                    table_name_use.append(i)
                column_name_use = []
                for i in column_names:
                    column_name_use.append(i)
                table_names_docx = []
                if doc == '.DS_Store' or '~' in doc:
                    continue
                file = root.folder + "/" + doc
                document = Document(file)
                for para in document.paragraphs:
                    if para.style.name == 'Detail - Heading Synopsis':
                        table_names_docx.append(para.text.strip())

                count = -1
                information_arr = []
                index_arr = []
                remove_arr = []
                for table_title in table_names:
                    count += 1
                    if table_title == '':
                        continue
                    table_num = table_names_docx.index(table_title)
                    table = document.tables[table_num]
                    data = [[cell.text for cell in row.cells] for row in table.rows]
                    df = pd.DataFrame(data)
                    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
                    try:
                        row_word = row_words[count][0]
                        addition = []
                        mask = np.column_stack([df[col].str.contains(row_word, na=False) for col in df])
                        find_result = np.where(mask == True)
                        if find_result:
                            try:
                                result = [find_result[0][0], find_result[1][0]]
                                addition.append(result)
                                index_arr.append(count)
                            except:
                                remove_arr.append(table_title)
                                continue
                            try:
                                result2 = [find_result[0][1], find_result[1][1]]
                                addition.append(result2)
                                information_arr.append(addition)
                            except:
                                information_arr.append(addition)

                    except:
                        continue
                iteration = -1
                for thing in remove_arr:
                    index = table_name_use.index(thing)
                    table_name_use.remove(thing)
                    del column_name_use[index]
                for table_title in table_name_use:
                    iteration += 1
                    if table_title == '':
                        continue
                    table_num = table_names_docx.index(table_title)
                    table = document.tables[table_num]
                    data = [[cell.text for cell in row.cells] for row in table.rows]
                    df = pd.DataFrame(data)
                    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
                    group = information_arr[iteration]
                    index_item = index_arr[iteration]
                    for piece in group:
                        row = piece[0]
                        get_cols = column_name_use[iteration][0].split(',')
                        column_label1 = get_cols[0].strip()
                        if len(get_cols) > 1:
                            column_label2 = get_cols[1].strip()
                        else:
                            column_label2 = 'Not existent'
                        if column_label1 in df:
                            output1 = df.loc[df.index[row], column_label1]
                            adding = output1.strip()
                            output1_list[index_item] = adding
                            if column_label2 in df:
                                output2 = df.loc[df.index[row], column_label2]
                                push = output2.strip()
                                output2_list[index_item] = push

                develop_sentences(output1_list, output2_list)


title = Label(root, text="Welcome to AutoSynopsis!", font=("Arial", 25))
title.pack()


conf_btn = Button(root, text="Get configuration file",
                  command=lambda: open_conf(my_label),
                  relief=SUNKEN, padx=20, pady=10,
                  font=("Arial", 16))
conf_btn.pack()

gen_syn_btn = Button(root, text="Generate Synopsis", command=lambda: gen_synopses(),
                     relief=SUNKEN, padx=20, pady=10,
                     font=("Arial", 16))
gen_syn_btn.pack()

start = Label(root, text="Please configure this before generating Synopses.", font=("Arial", 20))
start.pack()

root.mainloop()
