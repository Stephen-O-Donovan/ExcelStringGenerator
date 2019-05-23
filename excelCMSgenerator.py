
def createCMSString(parentName, levelName, parentIDcolumnLetter, parentNameColumnLetter,
                     parentInChildColumnLetter, firstCellNumber, lastCellNumber):

    fileName = ("%s.txt" % levelName)
    delimiter = 0.00

    if levelName == "Task":
        delimiter = 0.01
    elif levelName == "Activity":
        delimiter = 0.0001

    file = open(fileName, "w")
    taskString = '='
    for i in range(firstCellNumber, lastCellNumber):
        taskString += ("IF(%s%i=%s!$%s$%i,%s!$%s$%i+%0.6f,"
                       % (parentInChildColumnLetter, i, parentName,
                          parentNameColumnLetter, i, parentName, parentIDcolumnLetter, i, delimiter) )
    for i in range(firstCellNumber, lastCellNumber):
        taskString += ")"
    file.write(taskString+"\n")
    taskString = ''
    for cells in range(firstCellNumber+1, lastCellNumber):
        taskString += '='
        
        for i in range(firstCellNumber, lastCellNumber):
            taskString += ("IF(%s%i=%s!$%s$%i,%s!$%s$%i+%s%i-TRUNC(%s%i)+%0.6f,"
                           % (parentInChildColumnLetter, cells, parentName, parentNameColumnLetter, i,
                              parentName, parentIDcolumnLetter, i, parentIDcolumnLetter,
                              cells-1, parentIDcolumnLetter, cells-1, delimiter) )
        taskString += ('Select %s' % parentName)
        for i in range(firstCellNumber, lastCellNumber):
            taskString += ")"
        file.write(taskString+"\n")
        taskString = ''
    file.close()

parentName = "Task"
levelName = "Activity"
parentIDcolumnLetter = "A"
parentNameColumnLetter = "B"
parentInChildColumnLetter = "C"
firstCellNumber = 4
lastCellNumber = 20
createCMSString(parentName, levelName, parentIDcolumnLetter, parentNameColumnLetter,
                 parentInChildColumnLetter, firstCellNumber, lastCellNumber+1)

