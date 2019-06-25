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
    return [ int(tab[1]), col2num(tab[0]) ]

def isRowInCells( row, cells ):
    for item in cells:
        if item[0] == row:
            return 1
    return 0

def excel_style(row, col):
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    result = []
    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result) + str(row)

def excel2num(col):
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    result = 0
    l = list(col)
    l.reverse()
    cpt = 0
    for letter in l:
        result += (LETTERS.index(letter)+1)*(26**cpt)
        cpt += 1
    return result

# ToDo : gerer les cas A2:B5
def cellsExplodeTabs( cells ):
    new_cells = []
    for cell in cells:
        if ':' not in cell:
            new_cells.append(cell)
        else:
            # Should explode the table
            deb, fin = re.split(':', cell)
            col_deb = re.split('([^a-zA-Z]+)', deb)[0]
            col_fin   = re.split('([^a-zA-Z]+)', fin)[0]
            line_deb = col2tab(deb)[0]
            line_fin = col2tab(fin)[0]

            for line in range(line_deb, line_fin+1):
                for col in range( excel2num(col_deb), excel2num(col_fin)+1 ):
                    new_cells.append(excel_style(line, col))
    return new_cells

if __name__ == '__main__':
    cells = ["B13:D15"]
    print(cellsExplodeTabs(cells))
    exit(0)
    cell = "A4"
    print(col2tab(cell))
    col = 'A'
    print(excel2num(col))
    col = 'AA'
    print(excel2num(col))
    col = 'AAA'
    print(excel2num(col))
    col = 'B'
    print(excel2num(col))
    col = 'BC'
    print(excel2num(col))