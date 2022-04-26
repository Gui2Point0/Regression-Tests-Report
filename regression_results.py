import csv
import clean_csv_functions
import update_spreadsheet
import styling_excel

######################################################################################################

def ReportProcess(sortedAux, sortedCsv, version):
    counter = 0
    i = 0 
    finalInfoList = []
    # testNameList = []
    # testScoreList = []
    # testDurationList = []

    #Next Element
    for prevRow in sortedAux:
        if prevRow[0] == version:
            endTime = str(prevRow[5])[11:19] #Splits date and time
            counter += 1
            if counter > 1:
                #Element that is being added to the list
                row = sortedCsv[i]
                if row[0] == version:
                    startTime = str(row[5])[11:19] #Splits date and time
                    clean_csv_functions.CleanTestInfo(row)
                    row[1] = str(clean_csv_functions.GetTestRunTime(startTime, endTime))
                    
                    testName = row[0]
                    testDuration = row[1]
                    testScore = row[2]
                    
                    # testNameList.append(testName)
                    # testScoreList.append(testScore)
                    # testDurationList.append(testDuration)

                    finalInfoList.append([testName, testScore, testDuration])

                    i += 1

                else:
                    row.clear()
        else:
            prevRow.clear()
    
    return finalInfoList


def SetupCsvFiles(reportFileName):
    reportAux = open(reportFileName, 'r')
    report = open(reportFileName, 'r')

    csvreader = csv.reader(reportAux)
    csvreader2 = csv.reader(report)

    sortedAux = clean_csv_functions.SortCsvFile(csvreader)
    sortedCsv = clean_csv_functions.SortCsvFile(csvreader2)

    return sortedAux, sortedCsv

######################################################################################################

def StartProcess():
    reportFileName = str(input("\nEnter the report file name (e.g. 'report' - without extensions): ")) + '.csv'
    version = str(input("\nEnter the test build version (e.g. '12.0.7'): "))

    sortedAux, sortedCsv = SetupCsvFiles(reportFileName)
    testResultList = ReportProcess(sortedAux, sortedCsv, version)
    update_spreadsheet.OpenExcelFile(version, testResultList)
    styling_excel.SheetStyles()

######################################################################################################
        
StartProcess()