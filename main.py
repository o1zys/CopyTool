import os
import openpyxl.styles as sty
import openpyxl
import xlrd
import error
import global_var as gl


def copy_xls(old_xls):
    new_xls = openpyxl.Workbook()
    old_sheet = old_xls.sheet_by_index(0)
    new_sheet = new_xls.active
    max_row = old_sheet.nrows
    max_col = old_sheet.ncols

    for m in range(0, max_row):
        # for n in range(97, 97+max_col):
        for n in range(0, max_col):
            integer = n // 26
            remainder = n % 26
            if integer == 0:
                n_str = chr(remainder+97)
            else:
                n_str = chr(integer+96) + chr(remainder+97)

            i = '%s%d' % (n_str, m+1)
            cell = old_sheet.cell_value(m, n)
            new_sheet[i].value = cell
    return new_xls


def execute(file_a, file_b, v1, v2):
    output_dir = file_a.replace('.xlsx', '')
    # read output file, xls
    try:
        xls_a = xlrd.open_workbook(file_a)
    except Exception as e:
        error.Error.set_code(3, str(e))
        return
    try:
        xls_b = xlrd.open_workbook(file_b)
    except Exception as e:
        error.Error.set_code(3, str(e))
        return

    #old_sheet = old_xls.active
    old_sheet = xls_a.sheet_by_index(0)
    ref_sheet = xls_b.sheet_by_index(0)
    new_xls = copy_xls(xls_a)
    new_sheet = new_xls.active
    try:
        new_xls.save(output_dir + '_old.xlsx')
    except Exception as e:
        error.Error.set_code(6, str(e))
        return
    
    # construct dict shared id, dict value
    dict_sid = {}
    dict_val = {}
    dict_row = {}
    dict_lang_key = {}
    map_lang_col = {}
    tar_dict_sid = {}
    tar_dict_val = {}
    exist_row = []
    number_id = 0
    new_line = old_sheet.nrows
    for i in range(old_sheet.nrows):
        # read number id
        if i == 0:
            for j in range(old_sheet.ncols):
                if j > gl.col_langkey:
                    map_lang_col[old_sheet.cell_value(i, j)] = j
            continue
        if i < gl.trans_content_row:
            continue
        tmp_id = old_sheet.cell_value(i, gl.col_id)
        tmp_lang_key = old_sheet.cell_value(i, gl.col_langkey)
        dict_sid[tmp_id] = old_sheet.cell_value(i, gl.col_sid)
        dict_val[tmp_id] = tmp_lang_key
        dict_row[tmp_id] = i
        if tmp_lang_key != "":
            dict_lang_key[tmp_lang_key] = tmp_id

    # read ref_sheet
    tar_map_lang_col = {}
    for i in range(ref_sheet.nrows):
        if i == 0:
            for j in range(ref_sheet.ncols):
                if j > gl.col_langkey:
                    tar_map_lang_col[ref_sheet.cell_value(i, j)] = j
            continue
        if i < gl.trans_content_row:
            continue
        if ref_sheet.cell_value(i, gl.col_sid) != "":
            continue
        tar_id = ref_sheet.cell_value(i, gl.col_id)
        tar_val = ref_sheet.cell_value(i, gl.col_langkey)

        is_fill_content = False
        fill_type = sty.PatternFill(fill_type='solid', fgColor=gl.color_copy_modify)
        if v1 == 0 and dict_lang_key.__contains__(tar_val):
            fill_type = sty.PatternFill(fill_type='solid', fgColor=gl.color_copy_modify)
            tar_id = dict_lang_key[tar_val]
            for key, val in tar_map_lang_col.items():
                orig_content = old_sheet.cell_value(dict_row[tar_id], map_lang_col[key])
                tar_content = ref_sheet.cell_value(i, val)
                if tar_content != "" and tar_content != orig_content:
                    new_sheet.cell(dict_row[tar_id]+1, map_lang_col[key]+1, tar_content).fill = fill_type
            is_fill_content = True
        elif v1 == 1 and dict_val.__contains__(tar_id):
                if dict_val[tar_id] != tar_val and dict_sid[tar_id] == "":
                    fill_type = sty.PatternFill(fill_type='solid', fgColor=gl.color_copy_unique)
                    for key, val in tar_map_lang_col.items():
                        if key == "CNS":
                            continue
                        orig_content = old_sheet.cell_value(dict_row[tar_id], map_lang_col[key])
                        tar_content = ref_sheet.cell_value(i, val)
                        if tar_content != "" and tar_content != orig_content:
                            new_sheet.cell(dict_row[tar_id]+1, map_lang_col[key]+1, tar_content).fill = fill_type
                    is_fill_content = True

        if v2 == 0 and is_fill_content:
            orig_feature = old_sheet.cell_value(dict_row[tar_id], gl.col_feature)
            tar_feature = ref_sheet.cell_value(i, gl.col_feature)
            if tar_feature != orig_feature:
                new_sheet.cell(row=dict_row[tar_id]+1, column=gl.col_feature+1, value=tar_feature).fill \
                    = fill_type
            orig_term = old_sheet.cell_value(dict_row[tar_id], gl.col_term)
            tar_term = ref_sheet.cell_value(i, gl.col_term)
            if tar_term != orig_term:
                new_sheet.cell(row=dict_row[tar_id]+1, column=gl.col_term+1, value=tar_term).fill \
                    = fill_type
            orig_ignore = old_sheet.cell_value(dict_row[tar_id], gl.col_ignore)
            tar_ignore = ref_sheet.cell_value(i, gl.col_ignore)
            if tar_ignore != orig_ignore:
                new_sheet.cell(row=dict_row[tar_id]+1, column=gl.col_ignore+1, value=tar_ignore).fill \
                    = fill_type
            orig_desc = old_sheet.cell_value(dict_row[tar_id], gl.col_desc)
            tar_desc = ref_sheet.cell_value(i, gl.col_desc)
            if tar_desc != orig_desc:
                new_sheet.cell(row=dict_row[tar_id]+1, column=gl.col_desc+1, value=tar_desc).fill\
                    = fill_type
            orig_inst = old_sheet.cell_value(dict_row[tar_id], gl.col_instruction)
            tar_inst = ref_sheet.cell_value(i, gl.col_instruction)
            if tar_inst != orig_inst:
                new_sheet.cell(row=dict_row[tar_id]+1, column=gl.col_instruction+1, value=tar_inst).fill\
                    = fill_type

    try:
        os.remove(file_a)
    except Exception as e:
        error.Error.set_code(2, str(e))
        return

    try:
        new_xls.save(file_a)
    except Exception as e:
        error.Error.set_code(7, str(e))
        return

    new_xls.close()

