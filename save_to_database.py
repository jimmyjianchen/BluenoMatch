import xlrd
from xlwt import Workbook

# read from the calculated big table
path6 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Final_Matches.xls'
inputWorkbook6 = xlrd.open_workbook(path6)
inputWorksheet6 = inputWorkbook6.sheet_by_index(0)

# calculate row and column numbers
rowNumber6 = inputWorksheet6.nrows
columnNumber6 = inputWorksheet6.ncols


# read the file
path7 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/database.xls'
inputWorkbook7 = xlrd.open_workbook(path7)
inputWorksheet7 = inputWorkbook7.sheet_by_index(0)

# write the file
outputWorkbook7 = Workbook()
sheet7 = outputWorkbook7.add_sheet('Sheet 7')

# calculate row and column numbers
rowNumber7 = inputWorksheet7.nrows
columnNumber7 = inputWorksheet7.ncols

# write the file2
outputWorkbook7 = Workbook()
sheet7 = outputWorkbook7.add_sheet('Sheet 7')

everythingMatches = []
everythingDatabase = []
allPeopleMatches = []
allPeopleDatabase = []
extraNames = 0
matchDict = {}
databaseDict = {}

for i in range(rowNumber6):
    curr = []
    for j in range(columnNumber6):
        curr.append(inputWorksheet6.cell_value(i, j))
    everythingMatches.append(curr)

for i in range(rowNumber7):
    curr = []
    for j in range(columnNumber7):
        curr.append(inputWorksheet7.cell_value(i, j))
    everythingDatabase.append(curr)

for i in range(len(everythingMatches)):
    allPeopleMatches.append(everythingMatches[i][0])
    matchDict[everythingMatches[i][0]] = everythingMatches[i][1].split('\n')[0]

for i in range(len(everythingDatabase)):
    allPeopleDatabase.append(everythingDatabase[i][0])
    curr = []
    for j in range(len(everythingMatches[i])):
        if j != 0:
            if everythingDatabase[i][j] != '':
                curr.append(everythingDatabase[i][j])
    if not(matchDict[everythingDatabase[i][0]] in curr):
        if matchDict[everythingDatabase[i][0]] != '':
            curr.append(matchDict[everythingDatabase[i][0]])
    databaseDict[everythingDatabase[i][0]] = curr

for i in range(len(everythingDatabase)):
    sheet7.write(i, 0, everythingDatabase[i][0])
    for j in range(len(databaseDict[everythingDatabase[i][0]])):
        sheet7.write(i, j + 1, databaseDict[everythingDatabase[i][0]][j])

for i in range(len(allPeopleMatches)):
    if not(allPeopleMatches[i] in allPeopleDatabase):
        sheet7.write(rowNumber7 + extraNames, 0, allPeopleMatches[i])
        sheet7.write(rowNumber7 + extraNames, 1, everythingMatches[i][1].split('\n')[0])
        extraNames += 1

outputWorkbook7.save('database.xls')