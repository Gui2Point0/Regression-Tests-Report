import datetime as dt
import operator

from numpy import fix

def CleanTestInfo(row):
    del row[0] #Build Version
    del row[1] #OS
    del row[1] #Browser Info
    del row[1] #Resolution Info
    del row[3] #Test URL
    return row

def GetTestRunTime(start, end):
    start_dt = dt.datetime.strptime(start, '%H:%M:%S')
    end_dt = dt.datetime.strptime(end, '%H:%M:%S')

    runtime = end_dt - start_dt
    runtime -= dt.timedelta(seconds=5)

    return str(runtime)[2:]

def SortCsvFile(csv):
    sort = sorted(csv, key=operator.itemgetter(5))
    return sort