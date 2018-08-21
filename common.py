import json
import re
import os
from datetime import datetime

OUTPUT_FOLDER = "opendata"
INPUT_FOLDER = "finance_old"
COMPANY_STATUSES_FILE = "company_statuses.json"
NOT_SPACE = re.compile("\S+")
CURRENT_FILENAME = "filename.txt"
BRANCH_RE = re.compile("(\(((ВП).*|(філія).*?\)))")

def add_company_status(sheet_dict):
    with open(COMPANY_STATUSES_FILE, "r") as csf:
        status_dict = json.load(csf)
    sheet_dict['company_status'] = []
    for i in range(len(sheet_dict['company_code'])):
        try:
            sheet_dict["company_status"].append(status_dict[sheet_dict['company_code'][i]])
        except:
            sheet_dict["company_status"].append("")
    return sheet_dict

def refine_company_type(s):
    s = re.sub("\s+", " ", s)
    s = s.split("–")[0].strip()
    #s = s.replace("державні", "державне").replace("господарські", "господарське")
    #s = s.replace("підприємства", "підприємство").replace("товариства", "товариство")
    if "державн" and "підприємств" in s:
        s = "ДП"
    elif "господарськ" in s:
        if "понад" in s and "50" in s:
            s = "ГТ(б50)"
        elif "менше" in s and "50" in s:
            s = "ГТ(м50)"
    return s

def is_blank(cell):
    return NOT_SPACE.search(str(cell.value)) == None

def filename_part(d):
    QUARTERS_DICT = {"01": '12', "04": '3', "07": '6', "10": '9'}
    year  = d.year
    if d.month == 1:
        year = "_y" + str(year - 1)
    else:
        year = "_y" + str(year)
    month = datetime.strftime(d, '%m')
    return year + "_" + QUARTERS_DICT[month] + "m"

def extract_branch(s):
    branch_matched = BRANCH_RE.search(s)
    if branch_matched:
        return branch_matched.group(0)
    else:
        return ""
