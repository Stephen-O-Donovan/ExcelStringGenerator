"""
Generates strings to insert unique IDs into Excel
The numbers in the cells will auto generate and increment based
on the hierarchy of the item
Choose which level to generate cells for (Task/Activity/CheckList)
Choose the beginning cell and the number of cells
The script will generate a text file, simply copy and paste
the contents into the beginning cell in Excel
Excel will interpret all the functions for you, you
only need to copy and paste!

Author: Stephen O'Donovan
Date: 22/05/2019
Version: 2.0
"""

from tkinter import *

class GUI_Window:

    def __init__(self, master):

        mainframe = Frame(master)
        mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
        mainframe.columnconfigure(0, weight = 1)
        mainframe.rowconfigure(0, weight = 1)
        mainframe.pack(pady = 100, padx = 100)

        #Create the radio button list for item choice
        radioFrame = Frame(mainframe, borderwidth = 5, highlightthickness = 3, highlightbackground="black")
        radioFrame.grid(row = 1, column = 1)
        Label(radioFrame, text="Select the item type").pack()

        itemvar = StringVar()
        taskRadio = Radiobutton(radioFrame, text="Task", variable=itemvar, value="Task", indicatoron = 0, width=10)
        taskRadio.pack(anchor = W)

        activityRadio = Radiobutton(radioFrame, text="Activity", variable=itemvar, value="Activity", indicatoron = 0, width=10)
        activityRadio.pack(anchor = W)

        checkListRadio = Radiobutton(radioFrame, text="CheckList", variable=itemvar, value="CheckList", indicatoron = 0, width=10)
        checkListRadio.pack(anchor = W)

        itemvar.set("Task")

        #Create the entry fields
        entryFrame = Frame(mainframe, borderwidth = 5, highlightthickness = 3, highlightbackground="black")
        entryFrame.grid(column=3,row=1, sticky=(N,W,E,S) )
        Label(entryFrame, text="Enter the cell details").grid(row=3, column=3)

        pidColumnEntry = Entry(entryFrame, width=4)
        Label(entryFrame, text="Parent ID Cell Letter: ").grid(row=4, column=3)
        pidColumnEntry.grid(row=4, column=4)
        pidColumnEntry.delete(0, END)
        pidColumnEntry.insert(0, "A")

        pNameColumnEntry = Entry(entryFrame, width=4)
        Label(entryFrame, text="Parent Name Cell Letter: ").grid(row=5, column=3)
        pNameColumnEntry.grid(row=5, column=4)
        pNameColumnEntry.delete(0, END)
        pNameColumnEntry.insert(0, "B")

        pInChildColumnEntry = Entry(entryFrame, width=4)
        Label(entryFrame, text="Parent in Child Cell Letter: ").grid(row=6, column=3)
        pInChildColumnEntry.grid(row=6, column=4)
        pInChildColumnEntry.delete(0, END)
        pInChildColumnEntry.insert(0, "C")

        firstCellNumber = Entry(entryFrame, width=4)
        Label(entryFrame, text="First Cell Number: ").grid(row=7, column=3)
        firstCellNumber.grid(row=7, column=4)
        firstCellNumber.delete(0, END)
        firstCellNumber.insert(0, 4)

        lastCellNumber = Entry(entryFrame, width=4)
        Label(entryFrame, text="Last Cell Number: ").grid(row=8, column=3)
        lastCellNumber.grid(row=8, column=4)
        lastCellNumber.delete(0, END)
        lastCellNumber.insert(0, 50)


        # Create the generate and close buttons
        self.generateButton = Button(mainframe, text="Generate", command=lambda : self.createIDsInExcel(itemvar.get(), pidColumnEntry.get(), pNameColumnEntry.get(), pInChildColumnEntry.get(), int(firstCellNumber.get()), int(lastCellNumber.get())))
        self.generateButton.grid(row = 5, column = 1)

        self.closeButton = Button( mainframe, text="Close", fg="red", command=mainframe.quit)
        self.closeButton.grid(row = 5, column = 3)

    def createIDsInExcel(self, levelName, parentIDcolumnLetter, parentNameColumnLetter,
                     parentInChildColumnLetter, firstCellNumber, lastCellNumber):

        print("Generating text file")
        fileName = ("%s.txt" % levelName) #Change this to what and where you want the text file saved
        delimiter = 0.00
        multiplier = 1
        parentName = ''

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

root = Tk()

app = GUI_Window(root)

root.mainloop()
root.destroy()


