#261938489 Muhammad Mamoon 
# COMP 200 D FALL 23
# Assignment 1

# BASIC EXCEL SPREADSHEET
import sys

class Spreadsheet:
    def __init__(self):
        # Initializing variables in the constructor
        self.sheet = None
        self.rows = 0
        self.cols = 0
        self.cursor = [0, 0]
        self.selection = [None, None, None, None]
        self.prev = []

    def CreateSheet(self, rows, cols):
        if self.sheet != None:
            print("Error: Sheet already exists.")
            return

        # Using list comprehension to initialize a sheet with None value to allocate space
        self.sheet = [[None for n in range(cols)] for n2 in range(rows)]

        # Setting the number of rows and colums
        self.rows = rows
        self.cols = cols
        print("Sheet created.")
        

    def Goto(self, row, col):
        # Checking if row and column are within the sheet
        if row < self.rows and col < self.cols:
            self.cursor = [row-1, col-1]
        else:
            print("Invalid row or column.")

    def Insert(self, val):
        # self.savePrev()
        # CHecking if input is integer, then inserting the value at the selected place in cursor by indexing to get row and column
        if type(val) is int:
            self.sheet[self.cursor[0]][self.cursor[1]] = val
            self.savePrev()
        else:
            print("Please input an integer value.")

    def Delete(self):
        # Checking if a value exists at the specified cursor place, and if it does, deleting it by setting it to None
        if self.sheet[self.cursor[0]][self.cursor[1]] == None:
            print("Selected place already empty.")
        else:
            self.sheet[self.cursor[0]][self.cursor[1]] = None

    def ReadVal(self):
        # Same as above, checking if value exists, and printing it if it does
        if self.sheet[self.cursor[0]][self.cursor[1]] == None:
            print("No value available at the specified area to read.")
            return None
        else:
            value = self.sheet[self.cursor[0]][self.cursor[1]]
            print(f"The value at Position ({self.cursor[0] + 1, self.cursor[1] + 1}) is: {value}")
            return value

    def Select(self, row, col):
        # First again we check if row and column are within the sheet, and then setting our selection values
        if row < self.rows and col < self.cols:
                sr, sc = self.cursor[0], self.cursor[1]
                er, ec = row, col
                self.selection[0] = sr
                self.selection[1] = sc
                self.selection[2] = er
                self.selection[3] = ec
        else:
            print("Invalid row or column.")

    def GetSelection(self):
        # we convert the selection list into tuple and return it
        print(tuple(self.selection))
        return tuple(self.selection)

    def Sum(self, row, col):
        # Checking if selection isnt empty
        # self.savePrev()
        if self.selection[0] != None:
            totl = 0
            ur,uc,lr,lc = self.selection            # Setting values
            for rows in range(ur, lr + 1):          # Calculating sum by iterating over every value in the selection
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] != None:
                        totl += self.sheet[rows][cols]
            self.sheet[row-1][col-1] = totl             # Storing value at the specified target
            

        else:
            print("Invalid selection.")



    def Mul(self, row, col):
        # self.savePrev()
        # Similar to above implementation
        if self.selection[0] != None:
            totl = 1

            ur,uc,lr,lc = self.selection            # Setting values

            for rows in range(ur, lr + 1):          # Calculating product by iterating over every value in the selection
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] != None:
                        totl *= self.sheet[rows][cols]
            self.sheet[row-1][col-1] = totl             # Storing value at the specified target
            

        else:
            print("Invalid selection.")

    def Avg(self, row, col):
        # self.savePrev()
        # Modifying the implementation of sum to calculate average
        if self.selection[0] != None:
            totl = 0
            count = 0                               # Initializing count

            ur,uc,lr,lc = self.selection            # Setting values

            for rows in range(ur, lr + 1):          # Calculating sum by iterating over every value in the selection
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] != None:
                        totl += self.sheet[rows][cols]
                        count += 1

            if count > 0:
                avg = totl/count
                self.sheet[row-1][col-1] = int(avg)           # Storing value at the specified target (int() will make the avg non decimal and save space in the sheet)


        else:
            print("Invalid selection.")

    def Max(self, row, col):
        # Modifying sum function to accomodate max
        # self.savePrev()
        if self.selection[0] is not None and self.selection[2] is not None:
            maxv = None

            ur, uc, lr, lc = (
                min(self.selection[0], self.selection[2]),
                min(self.selection[1], self.selection[3]),
                max(self.selection[0], self.selection[2]),
                max(self.selection[1], self.selection[3]),
            )

            for rows in range(ur, lr + 1):
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] is not None:
                        if maxv is None or self.sheet[rows][cols] > maxv:
                            maxv = self.sheet[rows][cols]

            self.sheet[row-1][col-1] = maxv

            print(f"Max value is {maxv} between the selected area from Pos ({self.selection[0], self.selection[1]}) and Pos ({self.selection[2], self.selection[3]})")
        else:
            print("Invalid selection.")

    def PrintSheet(self):
        # NEED TO RE-WORK 17/10/23

        if self.sheet is not None:
            print("=========== SPREADSHEET ===========")
            n = 1
            for row in self.sheet:
                strng = f"{n} "
                for col in row:
                    if col is not None:
                        strng += "[%-4s]" % (str(col))      # Using string formatting to make all blocks equally sized (4 digits capacity)
        
                    else:
                        strng += "[%-4s]" % ("")
                n += 1
                print(strng)
            print("====================================")
        else:
            print("No spreadsheet found.")

    def Quit(self):
        print("Exiting spreadsheet.")
        sys.exit() # Using sys library to call the exit function

    # ------------------- END OF MAIN REQUIREMENTS, 17/10/23, BONUS TO BE WRITTEN AND TESTED LATER -------------


    # UNDO REDO AND SAVEPREV ADDED 23/10/23
    def savePrev(self):
        # Saving the state of the sheet before performing an action in self.prev
        # CREATING A DEEP COPY SO IT DOES NOT MATTER IF MAIN SHEET IS CHANGED LATER
        current = [row[:] for row in self.sheet]
        self.prev.append(current)


    def Undo(self):
        # From the array that I'm treating like a stack, we will pop the last available sheet before an action was performed and restore it
        if self.prev:
            # Removing most recent save to self.next var (This save is made after a function is called, so we have to remove it in order to access the form before that function was called)        
            self.next = self.prev.pop()

            last = self.prev.pop()

            '''print("last")
            print(last)'''

            self.sheet = last
            print("Sheet restored to its form before the last performed operation.")

    def Redo(self):
        # Quite simply, like we do in undo, instead of removing the most recent save, we will print it
        # Since that save is made AFTER the most recent function call, it will restore the function to the 'next' value
        if self.next:
            self.sheet = self.next
            print("Redone, sheet restored to the form AFTER the last function call.")

    def Save(self, file_name):
        try:
            # Using file handling we will write a new file, this means everytime a file is saved
            # the previous spreadsheet in that file will be overwritten
            # we will first store the num of rows and columns of our current spreadsgeet
            # then we'll iterate over every row's each column and store that data as it is in the file
            # separating data by /
            with open(file_name, 'w') as file:
                file.write(f"{self.rows}-{self.cols}\n") 
                for row in self.sheet:
                    for col in row:
                        if col != None:
                            file.write(f"{col}/")
                        else:
                            file.write("None/")
                    file.write('\n')
                    print("Spreadsheet saved to", file_name)
                file.write("\n\n\n")
        except FileNotFoundError:
            pass

    def Load(self, file_name):
        try:
            with open(file_name, 'r') as file:
                rc = file.readline().strip().split('-')
                # initializing sheet
                self.CreateSheet(int(rc[0]), int(rc[1]))
                rows = int(rc[0])
                cols = int(rc[1])

                # 20/10/23 2ND PUSH, NEED TO REWORK ON LOAD AND MAX()

                # Using nested loops to read the data, first we'll read the rows and splitting it will convert it into list
                # Then we'll iterate over every element of the row and fill the spreadsheet with that that element at that position
                for i in range(rows):
                    rdata = file.readline().strip().split("/")
                    for elm in range(len(rdata)):
                        if rdata[elm] == '':
                            rdata.pop(elm)
                    for j in range(cols):
                        cdata = rdata[j]
                        if cdata == "None":
                            self.sheet[i][j] = None
                        elif cdata == '':
                            pass
                        else:
                            self.sheet[i][j] = cdata
                print("Spreadsheet loaded successfully.")
        except FileNotFoundError:
            pass
                


def main():
    # creating spreadsheet w 6 rows and 6 columns
    spreadsheet = Spreadsheet()
    
    print("Welcome to Mamoon's python spreadsheet program")
    print("An automated spreadsheet will be created with 6 rows and 6 columns to test every function out.")
    spreadsheet.CreateSheet(6, 6)

    print("\n=====================================\n")

    # insert some values
    print("Inserting values")
    spreadsheet.Goto(2, 2)
    spreadsheet.Insert(10)
    spreadsheet.Goto(3, 3)
    spreadsheet.Insert(15)
    spreadsheet.Goto(4, 4)
    spreadsheet.Insert(20)
    spreadsheet.Goto(5,5)
    spreadsheet.Insert(344)
    spreadsheet.Goto(1,2)
    spreadsheet.Insert(12)
    spreadsheet.Goto(2,3)
    spreadsheet.Insert(23)
    spreadsheet.Goto(3,4)
    spreadsheet.Insert("abc")

    print("\n=====================================\n")
    print("Using readval function")

    # read values
    spreadsheet.Goto(2, 2)
    spreadsheet.ReadVal()
    spreadsheet.Goto(3, 3)
    spreadsheet.ReadVal()
    spreadsheet.Goto(4, 4)
    spreadsheet.ReadVal()

    print("\n=====================================\n")
    print("Calculating the sum of a selected area and reading the sum value which will be stored in a cell")
    # Create a selection rectangle and use sum
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(2, 4)
    spreadsheet.Sum(4, 2)
    spreadsheet.Goto(4, 2)
    spreadsheet.ReadVal()

    print("\n=====================================\n")
    print("Calculating the product of a selected area and reading the product value which will be stored in a cell")
    # Create a selection rectangle and use sum
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(2, 4)
    spreadsheet.Mul(4, 3)
    spreadsheet.Goto(4, 2)
    spreadsheet.ReadVal()

    print("\n=====================================\n")
    print("Calculating the average of selected area and reading the value stored in a cell")
    # Create a selection and use average
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(4, 4)
    spreadsheet.Avg(4, 1)
    spreadsheet.Goto(4, 1)
    spreadsheet.ReadVal()

    print("\n=====================================\n")
    print("Finding the max value from a selected area and reading that value as it will be stored in a cell")
    # Create another selection and finding the maximum
    spreadsheet.Goto(2, 1)
    spreadsheet.Select(4, 5)
    spreadsheet.Max(5, 1)
    spreadsheet.Goto(5, 1)
    spreadsheet.ReadVal()

    print("\n=====================================\n")
    print("Saving the data to a file and printing our sheet.\n")
    # Using file handling to save sheet and printing it
    spreadsheet.Save("sheets.txt")
    spreadsheet.PrintSheet()

    spreadsheet.Goto(1,1)
    spreadsheet.Insert(2)
    spreadsheet.savePrev()
    spreadsheet.Insert(3)
    spreadsheet.PrintSheet()
    spreadsheet.Undo()
    spreadsheet.PrintSheet()
    spreadsheet.Redo()
    spreadsheet.PrintSheet()


    print("\n=====================================\n")
    # Creating another spreadsheet now! to test out file handling
    spreadsheet2 = Spreadsheet()
    print("Another spreadsheet has now been initialized with 0 rows and 0 columns to test out file handling, this will read the data of the previous spreadsheet saved in the file and load it in itself.")
    spreadsheet2.Load("sheets.txt")
    spreadsheet2.PrintSheet()

main()

# 21/10/23 FINAL PUSH WITH MAIN IMPLEMENTATION. UNDO/REDO WILL BE ADDED IF TIME AVAILABLE