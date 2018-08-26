import os
import sys
import datetime
import common
import sheet0
import sheet1
import sheet2
import sheet3
import sheet4
import sheet5
import sheet6
import sheet7

files = os.listdir(common.INPUT_FOLDER)
files = [f for f in files if f.endswith(".xls")]

if not os.path.exists("finance.log"):
    f = open("finance.log", "w")
    f.close()
f = open("finance.log", "a")
sys.stdout = f
sys.stderr = f

print("----------------")
print(datetime.datetime.now())
for f in files:
    full_filename = os.path.join(common.INPUT_FOLDER, f)
    with open(common.CURRENT_FILENAME, "w") as cff:
        cff.write(full_filename)
    print(full_filename)
    sheet0.main()
    sheet1.main()
    sheet2.main()
    sheet3.main()
    sheet4.main()
    sheet5.main()
    sheet6.main()
    sheet7.main()
print("No errors were caught")
print("-----------------")

