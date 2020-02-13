##################################################
#   James Hunter Young
#   
#   ITAM Data Compile
#   
#   this code reads from a single ITAM AssetInventory Sheet
#   then places the values into another excel document
#   
#   
#   
#   
##################################################


#   ipmorts
import openpyxl, pprint

#   global variables
preWB = openpyxl.load_workbook('AssetInventory-PRE.xlsx')
preSheet = preWB['Computers']

postWB = openpyxl.load_workbook('AssetInventory-POST.xlsx')
postSheet = postWB['Computers']

def getData(sheetObj):
    assetData = []
    print('Opening workbook...')
    print('... grabbing data...')
    for row in range(2, sheetObj.max_row + 1):
        PCN = sheetObj.cell(row=row, column=2).value
        if PCN is None:
            # ignore those rows which have an empty PCN (assuming that it must be present)
            continue
        deviceType = sheetObj.cell(row=row, column=3).value
        deviceSN = sheetObj.cell(row=row, column=1).value
        userID = sheetObj.cell(row=row, column=6).value

        assetData.append([PCN, deviceType, deviceSN, userID])
    return assetData

def writePre(sheet_data):
    print('...writing data...')
    resultFile = open('crawl.py', 'w')
    resultFile.write('allData = ' + pprint.pformat(sheet_data))
    resultFile.close()
    print('...done.')


def main():
    preAssetData = getData(preSheet)
    postAssetData = getData(postSheet)
    print(preAssetData)
    print(postAssetData)
    writePre(preAssetData)

main()
