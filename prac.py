import tkinter as tk
from openpyxl import load_workbook
import datetime
def message():
    wb=load_workbook("students.xlsx")
    ws=wb["message"]
    from_id = input("Enter from id: ")
    to_id = input("Enter to_id: ")
    if from_id == "" or to_id == "":
        print("Enter both values")
    else:
        msg = input("Enter message: ")
        ws.append([from_id,to_id,msg])
        
        wb.save("students.xlsx")
message()
