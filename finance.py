import common
import os
import sheet0
import sheet1
import sheet2
import sheet3
import sheet4
import sheet5

files = os.listdir(common.INPUT_FOLDER)
files = [f for f in files if f.endswith(".xls")]
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

