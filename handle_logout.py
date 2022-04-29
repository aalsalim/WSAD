import os
import sys
import time
import ctypes
import json
import requests
import pyautogui
import pyodbc

from datetime import datetime
    

url = "https://leave-req.app/api/v1/wsadb"
 
headers = {"Content-Type": "application/json; charset=utf-8", "authorization": "choYeM744RLoq2ep0VT2hRC2h6NXIpm6E3yiQUwiDmUSDAhOzgONBO1R7Ylk"}

#######################################################################

value1 = os.environ["COMPUTERNAME"]
value2 = os.environ["USERNAME"]

#######################################################################

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
titles = []
def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append(buff.value)
    return True
EnumWindows(EnumWindowsProc(foreach_window), 0)
    
known_user = [i for i in titles if i.startswith("Information Window - Sectra IDS7")]
    
if known_user != []:
    value3 = known_user[0][35:]

else:
    value3 = "No Active User"

#######################################################################

value4 = pyautogui.position()

#######################################################################

now = datetime.now()
dt_string = now.strftime("%Y-%m-%d %H:%M:%S")

value5 = dt_string

#######################################################################

try:

    data = {"computer_name": value1, "windows_user_name": value2, "user_name": value3, "computer_status": "logout", "current_mouse_position": value4, "current_time": value5}

    #data = json.dumps({"computer_name": value1, "user_name": value2, "current_time": value3}, indent=4)


    response = requests.post(url, headers= headers, json= data)
     
    print("Status Code", response.status_code)
    print("JSON Response ", response.json())
    
except:
    print("Unhandled exception:", sys.exc_info()[0])


#######################################################################
#
#try:
#    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.221.44.28\radiology\PACSLO\wsdb.accdb;')
#    cursor = conn.cursor()
#
#
#    cursor.execute('insert into Table1 ([computer_name], [user_name], [current_time]) values (?,?,?)', (
#            value1, value2, value3
#        ))
#    cursor.commit()
#
#    print('Data Inserted')
#    
#except pyodbc.Error as e:
#    print("Error in connection", e)
