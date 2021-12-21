
import openpyxl
import os

print("hello world\n\n\n")

# set up "input.xlsx" for reading
input_file = "C:\\Users\\Peyton.Dexter\\Highstreet Work\\Resume Copy\\For upload\\input.xlsx"
wb = openpyxl.load_workbook(filename=input_file)
ws = wb.active

