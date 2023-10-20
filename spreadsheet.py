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
            self.cursor = [row, col]
        else:
            print("Invalid row or column.")

    def Insert(self, val):
        # CHecking if input is integer, then inserting the value at the selected place in cursor by indexing to get row and column
        if type(val) is int:
            self.sheet[self.cursor[0]][self.cursor[1]] = val
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
            print(value)
            return value

    def Select(self, row, col):
        # First again we check if row and column are within the sheet, and then setting our selection values
        if row < self.rows and col < self.cols:
                # Update the selection area correctly by finding the min and max values
                sr, sc = self.cursor[0], self.cursor[1]
                er, ec = row, col
                self.selection[0] = min(sr, er)
                self.selection[1] = min(sc, ec)
                self.selection[2] = max(sr, er)
                self.selection[3] = max(sc, ec)
        else:
            print("Invalid row or column.")

    def GetSelection(self):
        # we convert the selection list into tuple and return it
        print(tuple(self.selection))
        return tuple(self.selection)

    def Sum(self, row, col):
        # Checking if selection isnt empty
        if self.selection[0] != None:
            totl = 0
            ur,uc,lr,lc = self.selection            # Setting values
            for rows in range(ur, lr + 1):          # Calculating sum by iterating over every value in the selection
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] != None:
                        totl += self.sheet[rows][cols]
            self.sheet[row][col] = totl             # Storing value at the specified target
        else:
            print("Invalid selection.")



    def Mul(self, row, col):
        # Similar to above implementation
        if self.selection[0] != None:
            totl = 1

            ur,uc,lr,lc = self.selection            # Setting values

            for rows in range(ur, lr + 1):          # Calculating product by iterating over every value in the selection
                for cols in range(uc, lc + 1):
                    if self.sheet[rows][cols] != None:
                        totl *= self.sheet[rows][cols]
            self.sheet[row][col] = totl             # Storing value at the specified target
        else:
            print("Invalid selection.")

    def Avg(self, row, col):
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
                self.sheet[row][col] = int(avg)           # Storing value at the specified target (int() will make the avg non decimal and save space in the sheet)

        else:
            print("Invalid selection.")

    def Max(self, row, col):
        # Modifying sum function to accomodate max

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

            self.sheet[row][col] = maxv
            print(f"Max value is {maxv}")
        else:
            print("Invalid selection.")

    def PrintSheet(self):
        # NEED TO RE-WORK 17/10/23

        if self.sheet is not None:
            for row in self.sheet:
                strng = ""
                for col in row:
                    if col is not None:
                        strng += "[" + str(col) + "]"
                    else:
                        strng += "[ ]"
                print(strng)
        else:
            print("No spreadsheet found.")

    def Quit(self):
        print("Exiting spreadsheet.")
        sys.exit() # Using sys library to call the exit function

    # ------------------- END OF MAIN REQUIREMENTS, 17/10/23, BONUS TO BE WRITTEN AND TESTED LATER -------------

    def Undo(self):
        
        pass

    def Redo(self):
       
        pass

    def Save(self, file_name):
        try:
            with open(file_name, 'a') as file:
                file.write(f"{self.rows}-{self.cols}\n")
                for row in self.sheet:
                    for col in row:
                        if col is not None:
                            file.write(f"[{col}]")
                        else:
                            file.write("[ ]")
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
                for i in range(rows):
                    rdata = str(file.readline().strip())
                    data = rdata.split()
                    for j in range(cols):
                        if data[j] == "[ ]":
                            self.sheet[i][j] = "[ ]"
                        else:
                            self.sheet[i][j] = str(rdata[j])
                print("Spreadsheet loaded.")
        except FileNotFoundError:
            pass
                


def main():
    # Create a new spreadsheet with 5 rows and 5 columns
    spreadsheet = Spreadsheet()
    spreadsheet.CreateSheet(6, 6)
    print("Spreadsheet created with 6 rows and 6 columns.")

    # Move the cursor and insert some values
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
    # Read and print values
    spreadsheet.Goto(2, 2)
    spreadsheet.ReadVal()
    spreadsheet.Goto(3, 3)
    spreadsheet.ReadVal()
    spreadsheet.Goto(4, 4)
    spreadsheet.ReadVal()

    # Create a selection rectangle and calculate the sum
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(2, 4)
    spreadsheet.Sum(4, 2)
    spreadsheet.Goto(5, 5)
    spreadsheet.ReadVal()

    # Create a new selection and calculate the average
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(4, 4)
    spreadsheet.Avg(4, 1)
    spreadsheet.Goto(4, 4)
    spreadsheet.ReadVal()

    # Create another selection and find the maximum
    spreadsheet.Goto(2, 2)
    spreadsheet.Select(3, 5)
    spreadsheet.Max(5, 0)
    spreadsheet.Goto(4, 4)
    spreadsheet.ReadVal()

    spreadsheet.Save("sheets.txt")
    spreadsheet.PrintSheet()

s2 = Spreadsheet()
s2.Load("sheets.txt")
