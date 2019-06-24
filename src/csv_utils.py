import string
import re

def num2col(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def col2tab(col):
    tab = re.split('([^a-zA-Z]+)', col)
    return [ col2num(tab[0]), int(tab[1]) ]

def isRowInCells( row, cells ):
    for item in cells:
        if item[0] == row:
            return 1
    return 0

# ToDo : gerer les cas A2:B5
def cellsExplodeTabs( cells ):
    new_cells = []
    for cell in cells:
        if ':' not in cell:
            new_cells.append(cell)
        else:
            # Should explode the table
            deb, fin = re.split(':', cell)
            line = re.split('([^a-zA-Z]+)', deb)[0]
            deb = col2tab(deb)[1]
            fin = col2tab(fin)[1]
            for col in range(deb, fin+1):
                new_cells.append(line + str(col))
    return new_cells

if __name__ == '__main__':
    cell = "A24"
    print(col2tab(cell))