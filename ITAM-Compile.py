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

#postWB = openpyxl.load_workbook('AssetInventory-POST.xlsx')
#postSheet = postWB['Computers']


preAssetData = {}

#   copy the data from the pre into the master
def copyPre():
    print('Opening workbook...')
    print('... grabbing data...')
    for row in range(2,preSheet.max_row + 1):
        PCN = preSheet['B' + str(row)].value
        deviceType = preSheet['C' + str(row)].value
        deviceSN = preSheet['A' + str(row)].value
        userID = preSheet['F' + str(row)].value

        preAssetData.setdefault(PCN)
        preAssetData[PCN].setdefault(deviceType)

        preAssetData[PCN][deviceType][deviceSN][userID] += 1
    return preAssetData

def writePre():
    print('...writing data...')
    resultFile = open('crawl.py', 'w')
    resultFile.write('allData = ' + pprint.pformat(preAssetData))
    resultFile.close
    print('...done.')


def main():
    copyPre()
    writePre()

main()
