import configparser
import error
import re

# info
team = ""
project = ""
field_row = 0
key_row = 0
type_row = 0
content_row = 0
index_content_row = 0
trans_content_row = 4

# path
csv_dir = ""
index_file = ""
trans_file = ""
task_file = ""
glossary_file = ""
file_a = ""
file_b = ""
output_file = ""

# trans_col
col_id = 0
col_sid = 1
col_nid = 2
col_ignore = 3
col_hist = 4
col_designer = 5
col_sys = 6
col_feature = 7
col_term = 8
col_desc = 9
col_instruction = 10
col_langkey = 11

# index_col
index_csv_name = 0
index_csv_col = 1
index_designer = 2
index_sys = 3
index_feature = 4
index_term = 5

# constant_text_col
constant_text_designer = 2
constant_text_sys = 3
constant_text_feature = 4
constant_text_term = 5

# color
color_default = '00000000'
color_add = "ff0000"
color_modify = "ffff00"
color_fill = "33cc33"
color_copy_modify = '00b0f0'
color_copy_unique = '00ffff'

# constant
chs_pattern = "(\(|（)区别翻译\d*(\)|）)"
constant_text_file = "constanttext"
max_num_id_col = 1
array_string_pattern = "array<string>"
instruction_pattern = "\[(\w*)\]\[(\w*)\]"


def set_var_from_config():
    config = configparser.ConfigParser()
    config.read("config.txt", encoding='UTF-8-sig')

    global team
    team = config["DEFAULT"]["team"].replace(' ', '')
    global project
    project = config["DEFAULT"]["project"].replace(' ', '')

    global field_row
    field_row = int(config["file_row"]["field_row"].replace(' ', ''))
    global key_row
    key_row = int(config["file_row"]["key_row"].replace(' ', ''))
    global type_row
    type_row = int(config["file_row"]["type_row"].replace(' ', ''))
    global content_row
    content_row = int(config["file_row"]["content_row"].replace(' ', ''))
    global index_content_row
    index_content_row = int(config["file_row"]["index_content_row"].replace(' ', ''))
    global trans_content_row
    trans_content_row = int(config["file_row"]["trans_content_row"].replace(' ', ''))

    global csv_dir
    csv_dir = config["path"]["csv_dir"].replace(' ', '')
    global index_file
    index_file = config["path"]["index_file"].replace(' ', '')
    global trans_file
    trans_file = config["path"]["trans_file"].replace(' ', '').replace('$project', project)
    global task_file
    task_file = config["path"]["task_file"].replace(' ', '').replace('$project', project)
    global glossary_file
    glossary_file = config["path"]["glossary_file"].replace(' ', '').replace('$project', project)
    global file_a
    file_a = config["path"]["file_a"].replace(' ', '').replace('$project', project)
    global file_b
    file_b = config["path"]["file_b"].replace(' ', '').replace('$project', project)
    global output_file
    output_file = config["path"]["output_file"].replace(' ', '')

    global col_id
    col_id = int(config["trans_col"]["col_id"].replace(' ', ''))
    global col_sid
    col_sid = int(config["trans_col"]["col_sid"].replace(' ', ''))
    global col_nid
    col_nid = int(config["trans_col"]["col_nid"].replace(' ', ''))
    global col_ignore
    col_ignore = int(config["trans_col"]["col_ignore"].replace(' ', ''))
    global col_hist
    col_hist = int(config["trans_col"]["col_hist"].replace(' ', ''))
    global col_designer
    col_designer = int(config["trans_col"]["col_designer"].replace(' ', ''))
    global col_sys
    col_sys = int(config["trans_col"]["col_sys"].replace(' ', ''))
    global col_feature
    col_feature = int(config["trans_col"]["col_feature"].replace(' ', ''))
    global col_term
    col_term = int(config["trans_col"]["col_term"].replace(' ', ''))
    global col_desc
    col_desc = int(config["trans_col"]["col_desc"].replace(' ', ''))
    global col_instruction
    col_instruction = int(config["trans_col"]["col_instruction"].replace(' ', ''))
    global col_langkey
    col_langkey = int(config["trans_col"]["col_langkey"].replace(' ', ''))

    global index_csv_name
    index_csv_name = int(config["index_col"]["index_csv_name"].replace(' ', ''))
    global index_csv_col
    index_csv_col = int(config["index_col"]["index_csv_col"].replace(' ', ''))
    global index_designer
    index_designer = int(config["index_col"]["index_designer"].replace(' ', ''))
    global index_sys
    index_sys = int(config["index_col"]["index_sys"].replace(' ', ''))
    global index_feature
    index_feature = int(config["index_col"]["index_feature"].replace(' ', ''))
    global index_term
    index_term = int(config["index_col"]["index_term"].replace(' ', ''))

    global constant_text_designer
    constant_text_designer = int(config["constant_text_col"]["constant_text_designer"].replace(' ', ''))
    global constant_text_sys
    constant_text_sys = int(config["constant_text_col"]["constant_text_sys"].replace(' ', ''))
    global constant_text_feature
    constant_text_feature = int(config["constant_text_col"]["constant_text_feature"].replace(' ', ''))
    global constant_text_term
    constant_text_term = int(config["constant_text_col"]["constant_text_term"].replace(' ', ''))

    global color_add
    color_add = config["color"]["color_add"].replace(' ', '')
    global color_modify
    color_modify = config["color"]["color_modify"].replace(' ', '')
    global color_fill
    color_fill = config["color"]["color_fill"].replace(' ', '')
    global color_copy_modify
    color_copy_modify = config["color"]["color_copy_modify"].replace(' ', '')
    global color_copy_unique
    color_copy_unique = config["color"]["color_copy_unique"].replace(' ', '')
