
def createCMSString(levelName, parentIDcolumnLetter, parentNameColumnLetter,
                     parentInChildColumnLetter, firstCellNumber, lastCellNumber):

    fileName = ("%s.txt" % levelName)
    delimiter = 0.00
    multiplier = 1

    if levelName == "Task":
        parentName ="WP"
        delimiter = 0.01
    elif levelName == "Activity":
        parentName ="Task"
        delimiter = 0.0001
        multiplier = 100
    elif levelName == "CheckList":
        parentName ="Activity"
        delimiter = 0.000001
        multiplier = 10000

    file = open(fileName, "w")
    cmsString = '='
    for i in range(firstCellNumber, lastCellNumber):
        cmsString += ("IF(%s%i=%s!$%s$%i,%s!$%s$%i+%f,"
                       % (parentInChildColumnLetter, i, parentName,
                          parentNameColumnLetter, i, parentName, parentIDcolumnLetter, i, delimiter) )
    for i in range(firstCellNumber, lastCellNumber):
        cmsString += ")"
    file.write(cmsString+"\n")
    cmsString = ''
    for cells in range(firstCellNumber+1, lastCellNumber):
        cmsString += '='
        
        for i in range(firstCellNumber, lastCellNumber):
            cmsString += ("IF(%s%i=%s!$%s$%i,%s!$%s$%i+((%s%i*%i)-TRUNC(%s%i*%i))/%i+%f,"
                           % (parentInChildColumnLetter, cells, parentName, parentNameColumnLetter, i,
                              parentName, parentIDcolumnLetter, i, parentIDcolumnLetter,
                              cells-1, multiplier, parentIDcolumnLetter, cells-1, multiplier, multiplier, delimiter) )
        cmsString += ('Select %s' % parentName)
        for i in range(firstCellNumber, lastCellNumber):
            cmsString += ")"
        file.write(cmsString+"\n")
        cmsString = ''
    file.close()

levelName = "CheckList"
parentIDcolumnLetter = "A"
parentNameColumnLetter = "B"
parentInChildColumnLetter = "C"
firstCellNumber = 4
lastCellNumber = 20
createCMSString(levelName, parentIDcolumnLetter, parentNameColumnLetter,
                 parentInChildColumnLetter, firstCellNumber, lastCellNumber+1)

