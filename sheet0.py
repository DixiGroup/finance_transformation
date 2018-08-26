import csv
import xlrd
import xlsxwriter
import re
import json
import datetime
import common
import os

def load_workbook(wb):
    global sheet_dict, date_, company_type
    date_ = ""
    sheet = wb.sheet_by_index(SHEET_NUMBER)
    nrows = sheet.nrows
    for i in range(STARTING_ROW):
        if date_ == "":
            for j in range(sheet.ncols):
                if isinstance(sheet.cell(i, j).value, str):
                    date_re_matched = DATE_RE.search(sheet.cell(i, j).value)
                    if date_re_matched:
                        date_ = date_re_matched.group()
    for i in range(STARTING_ROW, nrows):
        if not (("господарські" in sheet.cell(i,1).value and  "товариства" in sheet.cell(i,1).value) or ("державн" in sheet.cell(i,1).value and  "підприємств" in sheet.cell(i,1).value)):
            for k in fields_dictionary.keys():
                sheet_dict[fields_dictionary[k]].append(strip_if_str(sheet.cell(i, int(k) - 1).value))
            non_working_status_type, non_working_status_decision = status_extract([sheet.cell(i, 10).value, sheet.cell(i, 11).value, sheet.cell(i, 12).value])
            sheet_dict["non_working_status_type"].append(non_working_status_type)
            sheet_dict["non_working_status_decision"].append(non_working_status_decision)
            sheet_dict['company_type'].append(company_type)
        else:
            company_type = common.refine_company_type(sheet.cell(i, 1).value)

def format_code(v):
    s = str(v).split('.')[0]    
    if len(s) < 8:
        s = "0" * (8 - len(s)) + s
    return s

def strip_if_str(v):
    if isinstance(v, str):
        v = v.strip()
    return v

def format_region_code(v):
    s = str(int(v)).strip()
    if s == "1":
        s = "0" + s
    extra_nulls = 10 - len(s)
    s = s + "0" * extra_nulls
    return s

def is_bold(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].bold == 1

  
def status_extract(statuses):
    global STATUSES
    statuses = [str(s).strip() for s in statuses]
    ind = [i for i in range(len(statuses)) if statuses[i] != ""]
    if len(ind) > 0:
        ind = ind[0]
        st = STATUSES[ind]
        if ind == 0:
            if statuses[ind].strip().startswith("РМ"):
                st += " - розпорядження майном"
            elif statuses[ind].strip().startswith("С"):
                st += " - санація"
            elif statuses[ind].strip().startswith("Л"):
                st += " - ліквідація"
        st_decision_matched = ST_DECISION_RE.search(statuses[ind])
        if st_decision_matched is not None:
            st_decision = st_decision_matched.group("status_decision")
        else:
            st_decision = ""
    else:
        st, st_decision = "", ""
    return st, st_decision

def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

def combine(dict_):
    new_dict = {}
    for i in range(len(dict_["company_code"])):
        new_dict[dict_["company_code"][i]] = dict_["is_working"][i]
        if len(dict_["non_working_status_type"][i]) > 0:
            new_dict[dict_["company_code"][i]] += "; " + dict_["non_working_status_type"][i]
        if len(dict_["add_info"][i]) > 0:
            new_dict[dict_["company_code"][i]] += "; " + dict_["add_info"][i]
    return new_dict

def is_working_extend(st):
    IS_WORKING_STATUSES = {"п": "працює", "нп": "не працює"}
    st = st.lower()
    if st in IS_WORKING_STATUSES.keys():
        return IS_WORKING_STATUSES[st]
    else:
        return st



def main():
    global STATUSES, SHEET_NUMBER, STARTING_ROW, DATE_RE, wb, fields_dictionary
    global ST_DECISION_RE, sheet_dict, date_, company_type

    OUTPUT_FILE = os.path.join(common.OUTPUT_FOLDER, 'companies_registry')
    DATE_FILE = "finance_date.txt"
    COMPANY_STATUSES_FILE = "company_statuses.json"
    VARIABLES_FILE = 'sheet0.csv'
    DIGITS_ROW_NUMBER = 5
    STARTING_ROW = 7
    SHEET_NUMBER = 0
    STATUSES = ["перебуває у процедурі банкрутства", "перебуває у процедурі реорганізації", "перебуває у процедурі у ліквідації за рішенням органу влади"]
    ST_DECISION_RE = re.compile("\((?P<status_decision>.*)\)")
    DATE_RE = re.compile("\d{2}\.\d{2}\.\d{4}")
    QUARTERS_DICT = {"01": 12, "04": 3, "07": 6, "10": 9}

    with open(VARIABLES_FILE, 'r') as vf:
        var_reader = csv.reader(vf)
        fields_dictionary = {}
        for l in var_reader:
            fields_dictionary[l[0]] = l[2]
    sheet_dict = {}
    headers = []
    keys_num = [int(k) for k in fields_dictionary.keys()]
    keys_num = sorted(keys_num)
    for k in keys_num:
        sheet_dict[fields_dictionary[str(k)]] = []
        headers.append(fields_dictionary[str(k)])
    headers = headers[:9] + ['non_working_status_type', 'non_working_status_decision'] + headers[9:]
    headers = headers[:3] + ["company_type"] + headers[3:]
    sheet_dict['non_working_status_type'] = []
    sheet_dict['non_working_status_decision'] = []
    sheet_dict['company_type'] = []
    company_type = ''
    with open(common.CURRENT_FILENAME, "r") as cff:
        input_file = cff.read()
    wb = xlrd.open_workbook(input_file, formatting_info = True)
    load_workbook(wb)
    sheet_dict['region_code'] = list(map(format_region_code, sheet_dict['region_code']))
    sheet_dict['company_code'] = list(map(format_code, sheet_dict['company_code']))
    sheet_dict['is_working'] = list(map(is_working_extend, sheet_dict['is_working']))
    statuses_dict = combine(sheet_dict)
    with open(COMPANY_STATUSES_FILE, "w") as csf:
        json.dump(statuses_dict, csf)
    with open(DATE_FILE, "w") as df:
        df.write(date_)
    date_datetime = datetime.datetime.strptime(date_, "%d.%m.%Y")
    headers = ["year", "period"] + headers
    if date_datetime.month == 1:
        sheet_dict["year"] = [int(date_[-4:]) - 1] * len(sheet_dict["company_code"])
    else:
        sheet_dict["year"] = [int(date_[-4:])] * len(sheet_dict["company_code"])
    sheet_dict["period"] = [QUARTERS_DICT[date_[3:5]]] * len(sheet_dict["company_code"])
    finance_list = dict_to_list(sheet_dict, headers)

    with open(OUTPUT_FILE + common.filename_part(date_datetime) + ".csv", "w", newline = "") as of:
        csvwriter = csv.writer(of)
        csvwriter.writerow(headers)
        for l in finance_list:
            if l[1] != "":
                line_to_write = l[:]
                #line_to_write[0] = datetime.datetime.strftime(line_to_write[0], "%d.%m.%Y")
                csvwriter.writerow(line_to_write)
                
    out_wb = xlsxwriter.Workbook(OUTPUT_FILE + common.filename_part(date_datetime) + ".xlsx")
    worksheet = out_wb.add_worksheet()
    datef = out_wb.add_format({'num_format':"dd.mm.yyyy"})
    numf = out_wb.add_format({'num_format':"0.00"})
    headerf = out_wb.add_format({'bold':True})
    for i in range(len(headers)):
        worksheet.write(0, i, headers[i], headerf)
    for i in range(len(finance_list)):
        for j in range(len(headers)):
            if j >  11 and j < 15:
                worksheet.write(i+1, j, finance_list[i][j], numf)
            else:
                worksheet.write(i+1, j, finance_list[i][j])
    out_wb.close()

