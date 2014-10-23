__author__ = 'Phoenix'

# from mmap import mmap,ACCESS_READ
from xlrd import open_workbook
from xlwt import Workbook,Style

# from GT
COOLSCI = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\For Jun (Cool Science Experiments).xlsx'
TOP10 = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\Top 10 Animal Fights.xlsx'
ASECOND = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\For Jun (A second a Day).xlsx'
#OA23012 = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\OA23012 Data (1).xlsx'
OA23012 = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\OA23012 Data (1).xlsx'
GTCOL = 6

# from DA
COOLSCI_DA = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\Cool Science Experiments You Can Do at Home.mp4_OA23000_New.xlsx'
TOP10_DA = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\Top 10 Animal Fights Cought by Camera.wmv_YAP23001_New.xlsx'
ASECOND_DA = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\A Second a Day.mp4_OA23000_New.xlsx'

TOP10_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\Top 10 Animal Fights Cought by Camera.wmv_OA23012.xlsx'
HOME_DEPOT_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\How To Prepare for a Painting Project - The Home Depot.wmv_OA23012.xlsx'
WORLD_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\What Would Happen If the World Lost Its.mp4_OA23012.xlsx'
EMBALMING_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\The Embalming Process.mp4_OA23012.xlsx'
DACOL = 5
isDA = True

xlsFileName = EMBALMING_23012_DA
# resultFile = 'E:\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\For Jun (Cool Science Experiments)_result.xls'
resultFile = xlsFileName.rpartition('.')[0] + '_result.xls'

wb = open_workbook(xlsFileName)

gpSheet = wb.sheet_by_index(0)

timeColIndex = DACOL

frameIntervals = []
preFrame = gpSheet.cell(2, timeColIndex).value
cFrame = 0

for ri in range(3, gpSheet.nrows, 1):
    cFrame = gpSheet.cell(ri, timeColIndex).value
    frameIntervals.append(round(cFrame - preFrame, 3))
    preFrame = cFrame
    # print gpSheet.cell(ri, 6).value


rwb = Workbook()
ws = rwb.add_sheet('intervals')
pre1 = frameIntervals[0]
pre2 = frameIntervals[1]
incRecord = []
decRecord = []
Record18 = []
Record15 = []
cValue = 0
ws.row(0).write(0, frameIntervals[0])
ws.row(1).write(0, frameIntervals[1])
for i in range(2, len(frameIntervals)-1):
    cValue = frameIntervals[i]
    if cValue == pre1 and pre1 == pre2:
        ws.row(i).write(0, cValue,Style.easyxf('pattern: pattern solid, fore_colour yellow'))
        incRecord.append(i)
    elif pre1 == 0.016 and cValue == pre1:
        ws.row(i).write(0, cValue,Style.easyxf('pattern: pattern solid, fore_colour red'))
        decRecord.append(i)
    elif cValue == 0.018:
        ws.row(i).write(0, cValue,Style.easyxf('pattern: pattern solid, fore_colour blue'))
        Record18.append(i)
    elif cValue == 0.015:
        ws.row(i).write(0, cValue,Style.easyxf('pattern: pattern solid, fore_colour brown'))
        Record15.append(i)
    else:
        ws.row(i).write(0, cValue)
    pre1 = pre2
    pre2 = cValue

for ind, val in enumerate(incRecord):
    ws.row(ind).write(1, val)

for ind, val in enumerate(decRecord):
    ws.row(ind).write(2, val)

for ind, val in enumerate(Record18):
    ws.row(ind).write(3, val)

for ind, val in enumerate(Record18):
    ws.row(ind).write(4, val)

ws.row(0).write(5, len(incRecord))
ws.row(1).write(5, len(decRecord))
ws.row(2).write(5, len(Record18))
ws.row(3).write(5, len(Record15))

rwb.save(resultFile)

# print gpSheet.cell(2, 6).value

# print gpSheet.nrows

# for sheet_index in range(wb.nsheets):
#     print sheet_index

# for s in wb.sheets():
#     print 'Sheet:', s.name
#     for row in range(s.nrows):
#         values = []
#         for col in range(s.ncols):
