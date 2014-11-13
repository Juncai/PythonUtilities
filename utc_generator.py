__author__ = 'Jon'


from xlrd import open_workbook
from xlwt import Workbook

# output file path
OUTPUT_PATH = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Comparison Oct 30\UTC.xls'

# reference file
REF_PATH = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\GT Issue Oct 22\GTVideo_Info.xlsx'

# embalming
# startUTC = 1401299816
# startUTCMSec = 904
# length = 1000

# read the results
wb = open_workbook(REF_PATH)
ref_sheet = wb.sheet_by_index(0)
start_utc_dict = {}

for ri in range(1, ref_sheet.nrows):
    start_utc_dict[str(int(ref_sheet.cell(ri, 1).value))] = (int(ref_sheet.cell(ri, 4).value), int(ref_sheet.cell(ri, 5).value), int(ref_sheet.cell(ri, 3).value))
    # print int(ref_sheet.cell(ri, 4).value)


outputFileName = OUTPUT_PATH
tsSequence = (17, 16, 17)
# write to the new xls file
rwb = Workbook()
for k in start_utc_dict:
    ws = rwb.add_sheet('UTC_' + str(k))
    cUTC = start_utc_dict[k][0]
    cUTCMSec = start_utc_dict[k][1]
    length = start_utc_dict[k][2]
    ws.row(0).write(0, cUTC)
    ws.row(0).write(1, cUTCMSec)
    for i in range(1, length):
        if cUTCMSec + tsSequence[(i - 1) % 3] >= 1000:
            cUTC += 1
        cUTCMSec = (cUTCMSec + tsSequence[(i - 1) % 3]) % 1000
        ws.row(i).write(0, cUTC)
        ws.row(i).write(1, cUTCMSec)

rwb.save(outputFileName)


