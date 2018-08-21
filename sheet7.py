import csv
import xlrd
import xlsxwriter
import common
import datetime
import os




def load_workbook(wb):
    global sheet_dict, company_type
    sheet = wb.sheet_by_index(SHEET_NUMBER)
    nrows = sheet.nrows
    for i in range(STARTING_ROW, nrows):
        if (not (common.is_blank(sheet.cell(i, 0)))) and (not is_company_type(sheet.cell(i, 0))):
            for k in fields_dictionary.keys():
                sheet_dict[fields_dictionary[k]].append(sheet.cell(i, int(k) - 1).value)
            sheet_dict['company_type'].append(company_type)
        elif is_company_type(sheet.cell(i, 0)):
            company_type = common.refine_company_type(sheet.cell(i, 1).value)            

def is_company_type(cell):
    if common.is_blank(cell):
        return False
    elif isinstance(cell.value, str) and "." in cell.value:
        return False
    elif isinstance(cell.value, str) and not ("." in cell.value):
        return True
    elif isinstance(cell.value, float) and int(cell.value) == cell.value:
        return True
    elif isinstance(cell.value, float) and int(cell.value) != cell.value:
        return False
    elif isinstance(cell.value, int):
        return True



def format_code(v):
    s = str(v).split('.')[0]    
    if len(s) < 8:
        s = "0" * (8 - len(s)) + s
    return s
  
def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

def main():
    global SHEET_NUMBER, STARTING_ROW, wb, fields_dictionary
    global sheet_dict, date_, company_type
    INPUT_FILE = "finance_old.xls"
    OUTPUT_FILE = os.path.join(common.OUTPUT_FOLDER, 'transfered_property')
    VARIABLES_FILE = 'sheet7.csv'
    DIGITS_ROW_NUMBER = 6
    STARTING_ROW = 8
    SHEET_NUMBER = 7
    DATE_FILE = "finance_date.txt"
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
    with open(common.CURRENT_FILENAME, "r") as cff:
        input_file = cff.read()
    wb = xlrd.open_workbook(input_file, formatting_info = True)
    sheet_dict["company_type"] = []
    company_type = ''
    load_workbook(wb)
    sheet_dict["company_code"] = list(map(format_code, sheet_dict["company_code"]))
    with open(DATE_FILE, "r") as df:
        date_ = df.read()
    date_datetime = datetime.datetime.strptime(date_, "%d.%m.%Y")
    if date_datetime.month == 1:
        sheet_dict["year"] = [int(date_[-4:]) - 1] * len(sheet_dict["company_code"])
    else:
        sheet_dict["year"] = [int(date_[-4:])] * len(sheet_dict["company_code"])
    sheet_dict["period"] = [QUARTERS_DICT[date_[3:5]]] * len(sheet_dict["company_code"])
    sheet_dict = common.add_company_status(sheet_dict)
    sheet_dict['branch'] = list(map(common.extract_branch, sheet_dict['company_name']))
    sheet_dict['company_name'] = [sheet_dict['company_name'][i].replace(sheet_dict['branch'][i],"").strip()for i in range(len(sheet_dict['company_name'])) if sheet_dict['branch'] != '' ]
    for i in range(len(sheet_dict['branch'])):
        if sheet_dict['branch'][i] != '':
                sheet_dict['branch'][i] = sheet_dict['branch'][i][1:-1] 
    headers = ['year', 'period'] + headers[:2] + ["company_type", "company_status", "branch"] + headers[2:]
    finance_list = dict_to_list(sheet_dict, headers)
    with open(OUTPUT_FILE + common.filename_part(date_datetime) + ".csv", "w") as of:
        csvwriter = csv.writer(of)
        csvwriter.writerow(headers)
        for l in finance_list:
            if l[0] != "":
                line_to_write = l[:]
                if isinstance(line_to_write[2], float):
                    d = xlrd.xldate_as_tuple(line_to_write[2],wb.datemode)
                    d = datetime.datetime(*d[0:6]).strftime("%d.%m.%Y")
                    line_to_write[2] = d
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
            if j >  6:
                worksheet.write(i+1, j, finance_list[i][j], numf)
            elif j == 6:
                worksheet.write(i+1, j, finance_list[i][j], datef)
            else:
                worksheet.write(i+1, j, finance_list[i][j])
    out_wb.close()
