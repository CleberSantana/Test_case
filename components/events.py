import os
import tkinter as tk
from tkinter import filedialog
import csv
import openpyxl
from openpyxl import load_workbook

l = []


def sheetlist(file_path):
    sheetlist1 = []
    wb = load_workbook(filename = os.path.basename(file_path))
    for i in wb.sheetnames:
        ws = wb[str(i)]
        if ws.sheet_properties.tabColor.value == 'FFA5A5A5':
            sheetlist1.append(i)
    wb.close()
    return sheetlist1


def clearlist(): # funtion to clean the KO list
    list = l.clear() # cleaning the list
    return list # returnin the list


def removedup(lista:list): # function for remove duplicated values of the list
    l = []
    for i in lista: # loop to iterate the list values
        if i not in l: # assign values to list if the value is not in it
            l.append(i) # append the value to the list
    return sorted(l) # sorting and returning the list


def ko(ct:int, flag:str): # function for create a list of KOs - 2 values assigned, 1st the number of spec, 2nd a flag to determine the function    
    if flag == '0': # comparing the flag value
        l.append(ct) # appeding the value "ct" to list
        list = removedup(l) # calling the funtion to remove a possible duplicated value
    elif flag == '1': # comparing the flag value
        list = clearlist() # calling the funtion to clean the list
    return list # returning the list 


def removevalue(ct:int): # funtion to remove the KO of the list
    l.remove(ct) # removing the spec of the list
    list = removedup(l) # calling the funtion to remove a possible duplicated value
    return list # returning the list


def remove_row(table, row): # funtion to remove rows from tables (lines in the file)
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr) # remove lines from file
   
   
def def_path(): # function for choose a file 
    root = tk.Tk() 
    root.withdraw()
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select file", filetypes=[("TestCase Files", "*.xlsx")])
    if file_path.endswith(".xlsx"):
        return os.path.abspath(file_path)
    else:
        return 'Select a file'
   
    
d = {}
def create_dict(title:int, dict1:dict): # function to create a dict 
    d[title] = dict1 # assing the spec and status to the dictionary
    return d # returning the dictionary


#creating list from cvs file
def filtering(file_path:str, currenttest:str):
    title = []
    d1 = {}
    d2 = {}
    c = 0
    dv = 0
    wb = load_workbook(filename = os.path.basename(file_path))
    ws = wb[str(currenttest)]
    totalrow = ws.max_row

    for row in ws.iter_rows(min_row = 5, max_row = totalrow+1, min_col = 3, max_col = 3):
        for cell in row:
            if "Test" in str(cell.value) or cell.row == totalrow:
                title.append(str(ws.cell(column = 5, row = cell.row).value))
                if c != 0:
                    d1 = d2.copy()
                    d = create_dict(title[dv], d1)
                    dv += 1
                    d2.clear()
            else:
                if ws.cell(column = 5, row = cell.row).value is None:
                    pass
                else:
                    d2[c] = {'step' : ws.cell(column = 5, row = cell.row).value,
                            'Expected' : ws.cell(column = 6, row = cell.row).value,
                            'Actual': ws.cell(column = 7, row = cell.row).value,
                            'sts':'',
                            'cellrow' : cell.row}
                    c += 1
    return d, title
                       
            
    

    '''for row in spamreader:
            dict1 = create_dict(c, row[0], row[1], row[2])
            if row[2] == 'KO':
                lista1 = ko(c,'0')
                print (lista1)
            else:
                lista1 = []
            c +=1
    return dict1, lista1 # returnig the dictionary'''