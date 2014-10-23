from xlrd import open_workbook
from xlwt import Workbook,Style

# from result read the intervals
TOP10 = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\Top 10 Animal Fights Cought by Camera.wmv_YAP23001_New_result.xls'
TOP10_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\Top 10 Animal Fights Cought by Camera.wmv_OA23012.xlsx'
HOME_DEPOT_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\How To Prepare for a Painting Project - The Home Depot.wmv_OA23012.xlsx'
WORLD_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\What Would Happen If the World Lost Its.mp4_OA23012.xlsx'
EMBALMING_23012_DA = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\Oct 14\The Embalming Process.mp4_OA23012.xlsx'

# output file path
OUTPUT_PATH = 'C:\Users\Jon\Dropbox\PlaIT Lab\eye_tracker\DataAnalysis\For Comparison\Top 10 Animal Fights Cought by Camera.wmv_YAP23001_UTC.xls'

# start time stamp
# startUTC = 1401300064
# startUTCMSec = 122

# home depot
# startUTC = 1405517341
# startUTCMSec = 485

# # world
# startUTC = 1405517748
# startUTCMSec = 864

# embalming
startUTC = 1401299816
startUTCMSec = 904

# read the results
xlsFileName = EMBALMING_23012_DA
xlsFileName = xlsFileName.rpartition('.')[0] + '_result.xls'
outputFileName = xlsFileName.rpartition('.')[0] + '_UTC.xls'
wb = open_workbook(xlsFileName)

gpSheet = wb.sheet_by_index(0)

intevalColIndex = 0

# initail variables

frameIntervals = []
# preFrame = gpSheet.cell(2, timeColIndex).value

for ri in range(0, gpSheet.nrows, 1):
    cFrame = gpSheet.cell(ri, intevalColIndex).value
    # print 1000 * cFrame
    frameIntervals.append(1000 * cFrame)


# write to the new xls file
rwb = Workbook()
ws = rwb.add_sheet('UTC')
ws.row(0).write(0, startUTC)
ws.row(0).write(1, startUTCMSec)
cUTCMSec = startUTCMSec
cUTC = startUTC

# wirte the first line

for i in range(0, len(frameIntervals), 1):
    if cUTCMSec + frameIntervals[i] >= 1000:
        cUTC += 1
    cUTCMSec = (cUTCMSec + frameIntervals[i]) % 1000
    ws.row(i+1).write(0, cUTC)
    ws.row(i+1).write(1, cUTCMSec)

rwb.save(outputFileName)

