import csv
import xlrd
import xlsxwriter
import re
import datetime
import common
import os

INPUT_FILE = "finance_old.xls"
OUTPUT_FILE = os.path.join(common.OUTPUT_FOLDER, 'financial_plans')
VARIABLES_FILE = 'sheet1.csv'
DATE_FILE = "finance_date.txt"
DIGITS_ROW_NUMBER = 6
STARTING_ROW = 7
SHEET_NUMBER = 1
DATE_RE = re.compile("\d{2}\.\d{2}\.\d{4}")
HEADERS = ["date", "company_name", "company_code", "company_type", "company_status", "fin_plan_adopted", "fin_plan_adoption_date", 'fin_plan_changes_dates', "sales_revenue_plan",   "sales_revenue_fact", "finance_result_plan", "finance_result_fact", "revenue_to_budget_plan", 
           "revenue_to_budget_fact", "revenue_to_state_shareholder_plan", "revenue_to_state_shareholder_fact", "capital_investments_plan", "capital_investments_fact", ]

def load_workbook(wb):
    global sheet_dict, company_type
    sheet = wb.sheet_by_index(SHEET_NUMBER)
    nrows = sheet.nrows
    for i in range(STARTING_ROW, nrows):
        if (not "Усього" in str(sheet.cell(i, 1).value)) and (not common.is_blank(sheet.cell(i,0))) and (not is_bold(sheet.cell(i, 1)) and (not common.is_blank(sheet.cell(i, 1)))):
            for k in fields_dictionary.keys():
                cell_value = sheet.cell(i, int(k) - 1).value
                if str(cell_value).strip() == "-":
                    cell_value = None
                sheet_dict[fields_dictionary[k]].append(cell_value)
            sheet_dict['company_type'].append(company_type)
        elif is_bold(sheet.cell(i, 1)) or ("господарські" in sheet.cell(i,1).value and  "товариства" in sheet.cell(i,1).value) or ("державн" in sheet.cell(i,1).value and  "підприємств" in sheet.cell(i,1).value):
            company_type = common.refine_company_type(sheet.cell(i, 1).value)

def is_bold(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].bold == 1

def format_code(v):
    s = str(v).split('.')[0]    
    if len(s) < 8:
        s = "0" * (8 - len(s)) + s
    return s

def plan_date_refine(cell):
    if isinstance(cell, str):
        if len(cell) > 0:
            return DATE_RE.findall(cell)
        else:
            return None
    else:
        d = xlrd.xldate_as_tuple(cell,0)
        d = [datetime.datetime(*d[0:6]).strftime("%d.%m.%Y")]
        return d
    
def extract_changes_date(x):
    if x:
        if len(x) > 1:
            return ";".join(x[1:])
        
def string_to_date(x):
    if x:
        return datetime.datetime.strptime(x[0], "%d.%m.%Y")
        
def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

def main():
    global SHEET_NUMBER, STARTING_ROW, DATE_RE, wb, fields_dictionary
    global sheet_dict, date_, company_type

    INPUT_FILE = "finance_old.xls"
    OUTPUT_FILE = os.path.join(common.OUTPUT_FOLDER, 'financial_plans')
    VARIABLES_FILE = 'sheet1.csv'
    DATE_FILE = "finance_date.txt"
    DIGITS_ROW_NUMBER = 6
    STARTING_ROW = 7
    SHEET_NUMBER = 1
    DATE_RE = re.compile("\d{2}\.\d{2}\.\d{4}")
    QUARTERS_DICT = {"01": 12, "04": 3, "07": 6, "10": 9}
    HEADERS = ["year", "period", "company_name", "company_code", "company_type", "company_status", "fin_plan_adopted", "fin_plan_adoption_date", 'fin_plan_changes_dates', "sales_revenue_plan",   "sales_revenue_fact", "finance_result_plan", "finance_result_fact", "revenue_to_budget_plan", 
            "revenue_to_budget_fact", "revenue_to_state_shareholder_plan", "revenue_to_state_shareholder_fact", "capital_investments_plan", "capital_investments_fact", ]


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
    with open(common.CURRENT_FILENAME, "r") as cff:
        input_file = cff.read()
    wb = xlrd.open_workbook(input_file, formatting_info = True)
    company_type = ''
    sheet_dict["company_type"] = []
    sheet_dict["company_status"] = []
    load_workbook(wb)
    sheet_dict["company_code"] = list(map(format_code, sheet_dict["company_code"]))
    sheet_dict["fin_plan_adoption_date"] = list(map(plan_date_refine, sheet_dict["fin_plan_adoption_date"]))
    sheet_dict['fin_plan_changes_dates'] = list(map(extract_changes_date, sheet_dict["fin_plan_adoption_date"]))
    sheet_dict["fin_plan_adoption_date"] = list(map(string_to_date, sheet_dict["fin_plan_adoption_date"]))
    sheet_dict['fin_plan_adopted'] = list(map(lambda s: s.lower().strip(), sheet_dict['fin_plan_adopted']))
    sheet_dict = common.add_company_status(sheet_dict)
    with open(DATE_FILE, "r") as df:
        date_ = df.read()
    date_datetime = datetime.datetime.strptime(date_, "%d.%m.%Y")
    headers = ["year", "period"] + headers
    if date_datetime.month == 1:
        sheet_dict["year"] = [int(date_[-4:]) - 1] * len(sheet_dict["company_code"])
    else:
        sheet_dict["year"] = [int(date_[-4:])] * len(sheet_dict["company_code"])
    sheet_dict["period"] = [QUARTERS_DICT[date_[3:5]]] * len(sheet_dict["company_code"])
    finance_list = dict_to_list(sheet_dict, HEADERS)
    with open(OUTPUT_FILE + common.filename_part(date_datetime) + ".csv", "w", newline="") as of:
        csvwriter = csv.writer(of)
        csvwriter.writerow(HEADERS)
        for l in finance_list:
            if l[0] != "":
                line_to_write = l[:]
                if line_to_write[7]:
                    line_to_write[7] = datetime.datetime.strftime(line_to_write[7],"%d.%m.%Y")
                csvwriter.writerow(line_to_write)
            
    out_wb = xlsxwriter.Workbook(OUTPUT_FILE + common.filename_part(date_datetime) + ".xlsx")
    worksheet = out_wb.add_worksheet()
    datef = out_wb.add_format({'num_format':"dd.mm.yyyy"})
    numf = out_wb.add_format({'num_format':"0.00"})
    headerf = out_wb.add_format({'bold':True})
    for i in range(len(HEADERS)):
        worksheet.write(0, i, HEADERS[i], headerf)
    for i in range(len(finance_list)):
        for j in range(len(HEADERS)):
            if finance_list[i][0] != "":
                if j == 7:
                    worksheet.write(i+1, j, finance_list[i][j], datef)
                elif j >  8:
                    worksheet.write(i+1, j, finance_list[i][j], numf)
                else:
                    worksheet.write(i+1, j, finance_list[i][j])
    out_wb.close()

