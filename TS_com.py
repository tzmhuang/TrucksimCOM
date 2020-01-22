import win32com.client
import win32process
import win32api
import win32gui
import win32con
import time
from pywinauto import Desktop
import threading

def call_back(hwnd,param):
    if win32process.GetWindowThreadProcessId(hwnd)[1] == param[0]:
        param[1].append(hwnd)


def get_handle(ts_hndl,class_name):
    handles = []
    process_id = win32process.GetWindowThreadProcessId(ts_hndl)[1]
    win32gui.EnumWindows(call_back,(process_id,handles))
    for h in handles:
        if win32gui.GetClassName(h) == class_name:
            return h
    raise Exception("Error! class_name not found!")
    #process_enum = win32process.EnumProcesses()    

def alt_key_combo(key_hex):
    alt_key = 0x12  #18
    win32api.keybd_event(alt_key,0,0,0) #Alt
    win32api.keybd_event(key_hex,0,0,0)
    win32api.keybd_event(key_hex,0,win32con.KEYEVENTF_KEYUP,0) #key release
    win32api.keybd_event(alt_key,0,win32con.KEYEVENTF_KEYUP,0)


ts_class = "CVIRTLVDChild00400000 68420000" #class name for Select Recent Databse window
ts_name = "Select Recent Database"  #window name
ts_name_2 = "License Settings"  #window name
TS_browse_clsName = "#32770"    #class name for 浏览文件夹 window
Lv1_tree = "\\桌面\\此电脑\\OS (C:)\\Truck_Sim"
Lv2_tree = "\\TruckSim2016.1_Data"
max_waiting_time = 20

def select_database_at_start():
    b_key = 0x42    #hexcode for 'b'
    s_key = 0x53    #hexcode for 's'
    hndl = win32gui.FindWindow(ts_class,ts_name)    #get handle for select recent database window
    win32gui.SetForegroundWindow(hndl)  #set select recent database to forground
    alt_key_combo(b_key)    #alt+b
    #tree
    time.sleep(1)   #waiting for 浏览文件夹 window to open
    diag_box_handle = get_handle(hndl,TS_browse_clsName)    #get handle of 浏览文件夹 window
    win = Desktop().window(handle = diag_box_handle)    #select 浏览文件夹 window using handle
    win.TreeView.GetItem(Lv1_tree).click()  #click on the Lv1_tree node
    win.TreeView.GetItem(Lv1_tree + Lv2_tree).click() #click on the Lv2_tree node
    win.button2.click() #press 确定
    counter = 0
    while not win32gui.FindWindow(ts_class,ts_name_2) and counter <max_waiting_time:
        time.sleep(1)
        counter+=1
    hndl_2 = win32gui.FindWindow(ts_class,ts_name_2)    #fetch handle for license window
    win32gui.SetForegroundWindow(hndl_2)    #set forground
    alt_key_combo(s_key)    #alt+s


def ts_run():   #TruckSim COM goes here
    h = win32com.client.Dispatch("TruckSim.Application")
    h.GoHome()
    ok = h.DataSetExists('','Baseline COM','External Control of Runs')
    assert ok == 1    
    h.DeleteDataSet('','New Run Made with COM','External Control of Runs')
    h.Gotolibrary('','Baseline COM','External Control of Runs')
    h.CreateNew()
    h.DatasetCategory('New Run Made with COM', 'External Control of Runs')
    h.Checkbox('#Checkbox8','1')
    h.Checkbox('#Checkbox3','1')
    h.Ring('#RingCtrl0','1')
    h.Yellow('*SPEED','30')
    h.Run('','')
    h.Checkbox('#Checkbox1','0')
    h.LaunchPlot()
    h.StartAnimator()

a = threading.Thread(target=ts_run) #creating a thread for TS COM
b = threading.Thread(target = select_database_at_start) #creatingg a thread for database selection

a.start()       #starting thread for TS COM
counter = 0
while not win32gui.FindWindow(ts_class,ts_name) and counter <max_waiting_time:
    time.sleep(1)
    counter+=1
print('here')
b.start()       #starting the thread for database selection
