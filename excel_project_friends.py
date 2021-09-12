import xlrd
from xlwt import Workbook
import sqlite3

# read the file
#path0 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Blueno Match Questionnaire week 1 (Responses).xls'
#inputWorkbook0 = xlrd.open_workbook(path0)
#inputWorksheet0 = inputWorkbook0.sheet_by_index(0)

# write the file
#outputWorkbook0 = Workbook()
#sheet0 = outputWorkbook0.add_sheet('Sheet 0')

# calculate row and column numbers
#rowNumber0 = inputWorksheet0.nrows
#columnNumber0 = inputWorksheet0.ncols

#everything = []
#allNames0 = []
#allGradeSelves0 = []
#allGradeMatches0 = []
#allLastTimes = []
#nameNumDict0 = {}
#numNameDict0 = {}

#numRowsWritten = 1

#for i in range(rowNumber0):
#    curr = []
#    for j in range(columnNumber0):
#        curr.append(inputWorksheet0.cell_value(i, j))
#    everything.append(curr)

# function that get rids of spaces in an input string (names)
def getRidOfSpaces(item):
    newItem = []
    for i in item:
        newI = i.strip()
        newItem.append(newI)
    return newItem

#def hasThisPerson0(name):
#    'function that checks if a given person has filled the survey'
#    for person in allNames0:
#        if person == name:
#            return True
#    return False

# check if two people are of matching types
#def areMatchTypes0(index1, index2):
#    gradeBool1 = allGradeSelves0[index1] in allGradeMatches0[index2].split(', ')
#    gradeBool2 = allGradeSelves0[index2] in allGradeMatches0[index1].split(', ')

#    if gradeBool1 and gradeBool2:
#        return True
#    else:
#        return False

#for i in range(rowNumber0):
#    if i != 0:
#        allNames0.append(everything[i][1])
#        nameNumDict0[allNames0[i - 1].strip()] = i - 1
#        numNameDict0[i - 1] = allNames0[i - 1].strip()
#        allGradeSelves0.append(everything[i][2])
#        allGradeMatches0.append(everything[i][3])
#        allLastTimes.append(everything[i][17])

#allNames0 = getRidOfSpaces(allNames0)

#for i in range(rowNumber0 - 1):
#    if allLastTimes[i] == "Yes, I didn't get a match last time.":
#        LastPriorityIndexes.append(i)

#for i in range(columnNumber0):
#    sheet0.write(0, i, inputWorksheet0.cell_value(0, i))

#for i in LastPriorityIndexes:
#    for j in range(columnNumber0):
#        sheet0.write(numRowsWritten, j, everything[i + 1][j])
#    numRowsWritten += 1

#for i in range(rowNumber0):
#    if i != 0:
#        if not(i in LastPriorityIndexes):
#            for j in range(columnNumber0):
#                if i + 1 < len(everything):
#                    sheet0.write(numRowsWritten, j, everything[i + 1][j])
#            numRowsWritten += 1

#outputWorkbook0.save('Survey_ranked_by_priority.xls')


# read the file
path1 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Blueno Match Questionnaire week 1 (Responses).xls'
inputWorkbook1 = xlrd.open_workbook(path1)
inputWorksheet1 = inputWorkbook1.sheet_by_index(0)

# write the file
outputWorkbook1 = Workbook()
sheet1 = outputWorkbook1.add_sheet('Sheet 1')

# calculate row and column numbers
rowNumber1 = inputWorksheet1.nrows
columnNumber1 = inputWorksheet1.ncols

# initiate the arrays that stores the values
allNames = []
allGradeSelves = []
allGradeMatches = []
allClasses = []
allSubFrees = []
allNativeLanguage = []
allHobbies = []
allHotTakes = []
allCulturalClubs = []
allClassesWeightings = []
allSubFreeWeightings = []
allNativeLanguageWeightings = []
allHobbiesWeightings = []
allHotTakesWeightigs = []
allCulturalClubsWeightings = []
allSuggestions = []


# dictionary that maps every person's name to their index in the arrays
nameNumDict = {}
numNameDict = {}

def hasThisPerson(name):
    'function that checks if a given person has filled the survey'
    for person in allNames:
        if person == name:
            return True
    return False

def updateWeightings(index):
    'When a person answers some answers to selected questions, their weighting to' 
    'this question should be automatically set to 0'
    
    if allHotTakesWeightigs[index] == 'I prefer not to answer this question':
        allHotTakesWeightigs[index] == 0

    if allCulturalClubs[index] == 'No' or allCulturalClubs[index] == '':
        allCulturalClubsWeightings[index] == 0

def calculateScore(index1, index2):
    'calculate the similarity score between two people'

    sum1 = classScore(index1, index2) * allClassesWeightings[index1] + substanceScore(index1, index2) * allSubFreeWeightings[index1] + nativeLanguageScore(index1, index2) * allNativeLanguageWeightings[index1] + hobbyScore(index1, index2) * allHobbiesWeightings[index1] + hotTakeScore(index1, index2) * allHotTakesWeightigs[index1] + culturalClubScore(index1, index2) * allCulturalClubsWeightings[index1]
    maximum1 = allClassesWeightings[index1] + allSubFreeWeightings[index1] + allNativeLanguageWeightings[index1] + allHobbiesWeightings[index1] + allHotTakesWeightigs[index1] + allCulturalClubsWeightings[index1]
    relative1 = sum1 / maximum1

    sum2 = classScore(index2, index1) * allClassesWeightings[index2] + substanceScore(index2, index1) * allSubFreeWeightings[index2] + nativeLanguageScore(index2, index1) * allNativeLanguageWeightings[index2] + hobbyScore(index2, index1) * allHobbiesWeightings[index2] + hotTakeScore(index2, index1) * allHotTakesWeightigs[index2] + culturalClubScore(index2, index1) * allCulturalClubsWeightings[index2]
    maximum2 = allClassesWeightings[index2] + allSubFreeWeightings[index2] + allNativeLanguageWeightings[index2] + allHobbiesWeightings[index2] + allHotTakesWeightigs[index2] + allCulturalClubsWeightings[index2]
    relative2 = sum2 / maximum2

    return (relative1 + relative2) / 2
 
def classScore(index1, index2):
    if allClasses[index1] == '':
        return 0

    if allClasses[index2] == '':
        return 0

    classes1 = allClasses[index1].split(', ')
    classes2 = allClasses[index2].split(', ')

    classes1Num = len(classes1)
    count = 0

    for class1 in classes1:
        if class1 in classes2:
            count += 1

    return count / classes1Num

def substanceScore(index1, index2):
    substance1 = allSubFrees[index1]
    substance2 = allSubFrees[index2]

    # if they both answer yes or no, return 1
    if substance1 == substance2:
        return 1

    # else return 0
    return 0

def nativeLanguageScore(index1, index2):
    if allNativeLanguage[index1] == '':
        return 0

    if allNativeLanguage[index2] == '':
        return 0

    nativeLanguages1 = allNativeLanguage[index1].split(', ')
    nativeLanguages2 = allNativeLanguage[index2].split(', ')

    nativeLanguages1Num = len(nativeLanguages1)
    count = 0

    for nativeLanguage1 in nativeLanguages1:
        if nativeLanguage1 in nativeLanguages2:
            count += 1

    return count / nativeLanguages1Num

def hobbyScore(index1, index2):
    if allHobbies[index1] == '':
        return 0

    if allHobbies[index2] == '':
        return 0

    hobbies1 = allHobbies[index1].split(', ')
    hobbies2 = allHobbies[index2].split(', ')

    hobbies1Num = len(hobbies1)
    count = 0

    for hobby1 in hobbies1:
        if hobby1 in hobbies2:
            count += 1

    return count / hobbies1Num

def hotTakeScore(index1, index2):
    hotTake1 = allHotTakes[index1]
    hotTake2 = allHotTakes[index2]
    
    if hotTake1 == 'I prefer not to answer this question':
        return 0

    num1 = 0
    num2 = 0

    if hotTake1 == 'Thats also what I think':
        num1 = 4
    elif hotTake1 == "I will try to see why they think that way and maybe accept their idea if I think they're right":
        num1 = 3
    elif hotTake1 == "It's none of my business":
        num1 = 2
    elif hotTake1 == 'I will try to talk about it with them and try to change their mind':
        num1 = 1

    if hotTake2 == "That's also what I think":
        num2 = 4
    elif hotTake2 == "I will try to see why they think that way and maybe accept their idea if I think they're right":
        num2 = 3
    elif hotTake2 == "It's none of my business":
        num2 = 2
    elif hotTake2 == 'I will try to talk about it with them and try to change their mind':
        num2 = 1

    return (4 - abs(num1 - num2)) / 4

def culturalClubScore(index1, index2):
    if allCulturalClubs[index1] == 'No' or allCulturalClubs[index1] == '':
        return 0

    if allCulturalClubs[index2] == 'No' or allCulturalClubs[index2] == '':
        return 0

    clubs1 = allCulturalClubs[index1].split(', ')
    clubs2 = allCulturalClubs[index2].split(', ')

    clubs1Num = len(clubs1)
    count = 0

    for club1 in clubs1:
        if club1 in clubs2:
            count += 1

    return count / clubs1Num

# check if two people are of matching types
def areMatchTypes(index1, index2):
    gradeBool1 = allGradeSelves[index1] in allGradeMatches[index2].split(', ')
    gradeBool2 = allGradeSelves[index2] in allGradeMatches[index1].split(', ')

    if gradeBool1 and gradeBool2:
        return True
    else:
        return False

# store values into the arrays
for i in range(rowNumber1):
    if i != 0:
        allNames.append(inputWorksheet1.cell_value(i, 1))
        nameNumDict[allNames[i - 1].strip()] = i - 1
        numNameDict[i - 1] = allNames[i - 1].strip()
        allGradeSelves.append(inputWorksheet1.cell_value(i, 2))
        allGradeMatches.append(inputWorksheet1.cell_value(i, 3))
        allClasses.append(inputWorksheet1.cell_value(i, 4))
        allSubFrees.append(inputWorksheet1.cell_value(i, 5))
        allNativeLanguage.append(inputWorksheet1.cell_value(i, 6))
        allHobbies.append(inputWorksheet1.cell_value(i, 7))
        allHotTakes.append(inputWorksheet1.cell_value(i, 8))
        allCulturalClubs.append(inputWorksheet1.cell_value(i, 9))
        allClassesWeightings.append(inputWorksheet1.cell_value(i, 10))
        allSubFreeWeightings.append(inputWorksheet1.cell_value(i, 11))
        allNativeLanguageWeightings.append(inputWorksheet1.cell_value(i, 12))
        allHobbiesWeightings.append(inputWorksheet1.cell_value(i, 13))
        allHotTakesWeightigs.append(inputWorksheet1.cell_value(i, 14))
        allCulturalClubsWeightings.append(inputWorksheet1.cell_value(i, 15))
        allSuggestions.append(inputWorksheet1.cell_value(i, 16))
        #allLastTimes.append(inputWorksheet1.cell_value(i, 17))

allNames = getRidOfSpaces(allNames)

# in some cases, the weightings should be updated
for i in range(rowNumber1 - 1):
    updateWeightings(i)

for i in range(rowNumber1 - 1):
        for j in range(rowNumber1 - 1):
            if i != j and areMatchTypes(i, j):
                score = calculateScore(i, j)
                sheet1.write(i + 1, j + 1, score)

for i in range(rowNumber1 - 1):
    sheet1.write(0, i + 1, allNames[i])
    sheet1.write(i + 1, 0, allNames[i])

outputWorkbook1.save('Calculated_Scores.xls')

# read from the calculated big table
path2 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Calculated_Scores.xls'
inputWorkbook2 = xlrd.open_workbook(path2)
inputWorksheet2 = inputWorkbook2.sheet_by_index(0)

# calculate row and column numbers
rowNumber2 = inputWorksheet2.nrows
columnNumber2 = inputWorksheet2.ncols

# write the file2
outputWorkbook2 = Workbook()
sheet2 = outputWorkbook2.add_sheet('Sheet 2')

for i in range(rowNumber2):
    if i != 0:
        count = 0
        for j in range(columnNumber2):
            if inputWorksheet2.cell_value(i, j) != '':
                count += 1
                if inputWorksheet2.cell_value(0, j) != '':
                    value = inputWorksheet2.cell_value(0, j) + '\n' + str(inputWorksheet2.cell_value(i, j))
                else:
                    value = inputWorksheet2.cell_value(i, j)
                sheet2.write(i - 1, count - 1, value)

outputWorkbook2.save('Only_Availables_Unsorted.xls')

# read from the calculated big table
path3 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Only_Availables_Unsorted.xls'
inputWorkbook3 = xlrd.open_workbook(path3)
inputWorksheet3 = inputWorkbook3.sheet_by_index(0)

# calculate row and column numbers
rowNumber3 = inputWorksheet3.nrows
columnNumber3 = inputWorksheet3.ncols

# write the file2
outputWorkbook3 = Workbook()
sheet3 = outputWorkbook3.add_sheet('Sheet 3')

def scoreToName(list1, list2):
    newList = []
    for i in list1:
        for j in list2:
            splitted = j.split('\n')
            if len(splitted) == 2:
                if i == splitted[1]:
                    newList.append(splitted[0] + '\n' + str(i))
    return newList

# an array with all scores
scoresArray = []
for i in range(rowNumber3):
    currscoreArray = []
    originalArray = []
    for j in range(columnNumber3):
        if j != 0:
            originalArray.append(inputWorksheet3.cell_value(i, j))
            values = inputWorksheet3.cell_value(i, j).split('\n')
            if values != ['']:
                values2 = values[1]
                currscoreArray.append(values2)
    currscoreArray.sort(reverse = True)
    scoresArray.append(scoreToName(currscoreArray, originalArray))

for i in range(rowNumber3):
    for j in range(columnNumber3):
        if j == 0:
            sheet3.write(i, j, inputWorksheet3.cell_value(i, j))
        else:
            if j - 1 < len(scoresArray[i]):
                sheet3.write(i, j, scoresArray[i][j - 1])
outputWorkbook3.save('Only_Availables_Sorted.xls')

# read from the calculated big table
path4 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Only_Availables_Sorted.xls'
inputWorkbook4 = xlrd.open_workbook(path4)
inputWorksheet4 = inputWorkbook4.sheet_by_index(0)

# calculate row and column numbers
rowNumber4 = inputWorksheet4.nrows
columnNumber4 = inputWorksheet4.ncols

# write the file2
outputWorkbook4 = Workbook()
sheet4 = outputWorkbook4.add_sheet('Sheet 4')

chosen = []
chosenPlus = []

inputWorksheetArray = []
for i in range(rowNumber4):
    curr = []
    for j in range(columnNumber4):
        curr.append(inputWorksheet4.cell_value(i, j))
    inputWorksheetArray.append(curr)

for i in range(rowNumber4):
    if not(inputWorksheetArray[i][0] in chosen):
        number = 1
        if len(inputWorksheetArray[i]) > 1:
            while number < len(inputWorksheetArray[i]) - 1 and inputWorksheetArray[i][number].split('\n')[0] in chosen:
              number += 1
            chosen.append(inputWorksheetArray[i][number].split('\n')[0])
            chosen.append(inputWorksheetArray[i][0])
            chosenPlus.append(inputWorksheetArray[i][number])
        else:
            chosen.append('')
            chosen.append(inputWorksheetArray[i][0])
            chosenPlus.append('')
    else:
        chosen.append('')
        chosen.append(inputWorksheetArray[i][0])
        chosenPlus.append('')

for i in range(rowNumber4):
        sheet4.write(i, 0, inputWorksheet4.cell_value(i, 0))
        if i < len(chosenPlus):
            sheet4.write(i, 1, chosenPlus[i])

outputWorkbook4.save('Final_Matches_Half.xls')

# read from the calculated big table
path5 = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Final_Matches_Half.xls'
inputWorkbook5 = xlrd.open_workbook(path5)
inputWorksheet5 = inputWorkbook5.sheet_by_index(0)

# calculate row and column numbers
rowNumber5 = inputWorksheet5.nrows
columnNumber5 = inputWorksheet5.ncols

# write the file2
outputWorkbook5 = Workbook()
sheet5 = outputWorkbook5.add_sheet('Sheet 5')

inputWorksheetArray5 = []
for i in range(rowNumber5):
    curr = []
    for j in range(columnNumber5):
        curr.append(inputWorksheet5.cell_value(i, j))
    inputWorksheetArray5.append(curr)

hashmap = {}
allKeys = []
for i in range(rowNumber5):
    if len(inputWorksheetArray5[i]) > 1 and len(inputWorksheetArray5[i][1].split('\n')) > 1:
        value = inputWorksheetArray5[i][0] + '\n' + inputWorksheetArray5[i][1].split('\n')[1]
        hashmap[inputWorksheetArray5[i][1].split('\n')[0]] = value
        allKeys.append(inputWorksheetArray5[i][1].split('\n')[0])

for i in range(rowNumber5):
    sheet5.write(i, 0, inputWorksheetArray5[i][0])
    if inputWorksheet5.cell_value(i, 1) != '':
        sheet5.write(i, 1, inputWorksheetArray5[i][1])
    elif inputWorksheetArray5[i][0] in allKeys:
        sheet5.write(i, 1, hashmap[inputWorksheetArray5[i][0]])
    else:
        sheet5.write(i, 1, '')

outputWorkbook5.save('Final_Matches.xls')