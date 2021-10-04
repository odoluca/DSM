print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__,__name__,str(__package__)))

"""--------------------BEGIN: COM OBJECTS-----------------------"""
from pythoncom import GetRunningObjectTable, CreateBindCtx
import sys
import threading
import trace
import ctypes
import time
from difflib import SequenceMatcher
from validator_collection import is_url
import string
import win32gui,win32process,win32con,win32com
import win32com.client as win32client
# import pywinauto #DOESNT WORK WITH PYINSTALLER
import pyautogui
import pyperclip
import psutil
from time import sleep
import os
from difflib import SequenceMatcher
import datetime
import pickle
import win32api,win32ui
import subprocess

IGNORED_FILE_EXTENSIONS=["xlam","dotm"] #files with these extensions are not reopened even if detected during session save

MAX_CPU_LOAD=10
CLOSE_UNWANTED=False
IGNORE_MINIMIZED=False

PROGRAM_ICO = "program.ico"
CHECK_ICO = "check.ico"
PROGRAM_TITLE = "Desktop Session Manager"

class com_object_info:
    associated=False

    def __init__(self,full_address,class_id,hash,system_moniker):
        self.full_address,self.class_id,self.hash,self.system_moniker=GetAbsolutePath(full_address),class_id,hash,(system_moniker==4)
        print("full adress",full_address)
        self.is_possible_work_file=IsPossibleWorkingFile(self.full_address)
        if self.is_possible_work_file:
            self.filename=GetFilenameFromFullAddress(self.full_address)

    def __repr__(self):
        return "com object info - Address: "+ str(self.full_address+" - Class ID:"+str(self.class_id)+" - System Moniker:"+str(self.system_moniker) )

    @staticmethod
    def FindRunningComObjectsAsInfo(ignore_system_moniker=True,only_workable_files=True):
        running_coms=GetRunningObjectTable()
        context=CreateBindCtx(0)
        monikers=running_coms.EnumRunning()
        moniker_infos=[]

        for moniker in monikers:
            moniker_info=(moniker.GetDisplayName(context,moniker),moniker.GetClassID(),moniker.Hash(),moniker.IsSystemMoniker())
            if ignore_system_moniker and moniker_info[3]==4: #if it is a system moniker returns 4, then...skip.
                continue
            c=com_object_info(*moniker_info)
            if only_workable_files and c.HasPossibleWorkingFile():
                moniker_infos.append(c)
        return moniker_infos

    @staticmethod
    def FindRunningComObjects(ignore_system_moniker=True):
        running_coms=GetRunningObjectTable()
        context=CreateBindCtx(0)
        monikers=running_coms.EnumRunning()
        monikers=[x for x in monikers if (x.IsSystemMoniker()!=4 or ignore_system_moniker==False)]
        return monikers

    def HasPossibleWorkingFile(self):
        if type(self.full_address) != str:
            raise Exception(
            "IsPossibleWorkingFile() requires string for file_address parameter. Maybe you passed a tuple?")
            # return False
        # if address.contains("!"): #case with non system monikers.
        # if starts with http we still allow as it can be through OneDrive
        file_extension = self.full_address.split(".")[-1].lower()
        if file_extension in IGNORED_FILE_EXTENSIONS and os.path.isfile(self.full_address):
            return False
        else:
            return True

"""----------------MISC FUNCTIONS----------------"""
def IsPossibleWorkingFile(file_address):
    if type(file_address)!= str:
        raise Exception("IsPossibleWorkingFile() requires string for file_address parameter. Maybe you passed a tuple?", type(file_address))
        # return False
    # if address.contains("!"): #case with non system monikers.
    # if starts with http we still allow as it can be through OneDrive
    file_extension=(file_address.split(".")[-1].lower())

    if file_extension not in IGNORED_FILE_EXTENSIONS and (CheckFileExists(file_address) or is_url(file_address) ):
        return True
    else:
        return False

def CheckFileExists(file_address):
    return os.path.isfile(file_address), "file do not exist!"


def GetAbsolutePath(file_address):
    if os.path.isfile(file_address):
        return os.path.abspath(file_address)
    elif is_url(file_address):
        return file_address
    else:
        return ""

def GetFilenameFromFullAddress(file_address,ignore_nonstring=False):
    if ignore_nonstring and type(file_address)!=str:
        return ""
    r=file_address.split("\\")[-1]
    return r

def RemovePuctuations(s):
    s.translate(str.maketrans('', '', string.punctuation+string.whitespace+" "))
    return s


def wait_cpu_usage_lower(threshold=None, timeout=None, check_interval=0.02, initial_sleep=0.02):
    """Wait until process CPU usage percentage is less than the specified threshold"""
    if threshold==None:
        global MAX_CPU_LOAD
        threshold=MAX_CPU_LOAD
    total_time=0
    sleep(initial_sleep)
    while(psutil.cpu_percent()>threshold):

        sleep(check_interval)
        total_time+=check_interval
        if timeout!=None and total_time>timeout:
            return "timeout"
    return "cpu"
    #TODO: consider cpu count for threshold decision


def PopMultipleFromList(lst,idxs):
    return [lst[x] for x in range(len(lst)) if x not in idxs]



"""--------------------END: MISC FUNCTIONS-----------------------"""
"""--------------------END: COM OBJECTS-----------------------"""
"""
https://pbpython.com/windows-com.html

"""
"""--------------------BEGIN: BROWSERS-----------------------"""
BROWSERS=["chrome.exe","firefox.exe","explorer.exe"] #May change to exe file name, may be useful for explorer

class Browser():
    MAXTABNO=50
    def __init__(self,name,bar_select=("alt","d"),tab_switch=("ctrl","tab")):
        self.name=name
        self.tab_switch=tab_switch
        self.bar_select=bar_select

    @staticmethod
    def GetThisTabUrls(wnd):
        # fm=fullscreen_message("Please wait...") #causes systray to close

        all_urls=[]
        isBrowser=False
        for br in BROWSERS:
            if GetFilenameFromFullAddress(wnd.exe)==br:
                BR=Browser(br)
                if (not wnd.is_minimized()) :  # only focus on nonminimized windows. Maybe this should be optional???
                    print("--->", wnd)

                    # # region BRING WINDOW FOREGROUND
                    # while (
                    #         win32gui.GetForegroundWindow() != wnd.hwnd):  # makesure the target window is in the foreground
                    #     print("fore:", win32gui.GetForegroundWindow(), "wanted:", wnd.hwnd)
                    #     win32gui.ShowWindow(wnd.hwnd, win32con.SW_MINIMIZE)
                    #     wait_cpu_usage_lower()
                    #
                    #     if (win32gui.GetForegroundWindow() != wnd.hwnd):
                    #         win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0,
                    #                              0)  # necassary to avoid error...Dont as why, Dont ask how I found it, I just felt it should work.
                    #         win32api.SetCursorPos((-10000, 500))
                    #         if wnd.is_minimized():
                    #             if wnd.was_maximized():
                    #                 win32gui.ShowWindow(wnd.hwnd, win32con.SW_MAXIMIZE)
                    #             else:
                    #                 win32gui.ShowWindow(wnd.hwnd, win32con.SW_RESTORE)
                    #         else:
                    #             win32gui.ShowWindow(wnd.hwnd, win32con.SW_SHOW)
                    #         wait_cpu_usage_lower()
                    #         win32gui.SetForegroundWindow(wnd.hwnd)
                    #     sleep(1)
                    #     wait_cpu_usage_lower()
                    #     # win32gui.SetFocus(wnd.hwnd) #access denied?
                    #     # wait_cpu_usage_lower()
                    #     # chrome_window.minimize()
                    #     # chrome_app.wait_cpu_usage_lower()  # necessary
                    #     # chrome_window.restore()
                    #     # chrome_window.set_focus()
                    #     # chrome_app.wait_cpu_usage_lower()  # necessary
                    #     # sleep(1) #necessary
                    #     # win32gui.ShowWindow(wnd.hwnd,win32con.SW_NORMAL)
                    #     # win32gui.BringWindowToTop(wnd.hwnd)
                    #     # win32gui.SetForegroundWindow(wnd.hwnd)
                    #     tries = 0
                    #     while (win32gui.GetForegroundWindow() != wnd.hwnd):
                    #         # win32gui.ShowWindow(win32gui.GetForegroundWindow(),win32con.SW_FORCEMINIMIZE)
                    #         # win32api.SetCursorPos((-10000, 500))
                    #         # win32gui.ShowWindow(wnd.hwnd, win32con.SW_SHOW)
                    #         # win32gui.BringWindowToTop(wnd.hwnd)
                    #         tries += 1
                    #         wait_cpu_usage_lower()
                    #         if tries > 3:
                    #             raise Exception("Cannot focus on window " + str(wnd.hwnd))
                    # # win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE) #minimize this window and try again
                    # # chrome_window.set_focus()
                    #
                    # # endregion
                    # # chrome_app.wait_cpu_usage_lower()  # necessary
                    # wait_cpu_usage_lower()
                    # # sleep(0.5)
                    window.FocusWindow(wnd.hwnd)

                    loop_complete = False
                    total_tabs=0
                    tab_titles = []
                    tab_urls = []
                    while (not loop_complete):
                        if total_tabs>Browser.MAXTABNO:
                            break
                        # tab_titles.append(chrome_window.element_info.name)
                        # win32api.keybd_event()
                        pyautogui.hotkey(*BR.bar_select)
                        # chrome_app.wait_cpu_usage_lower()  # necessary
                        wait_cpu_usage_lower()
                        pyautogui.hotkey("ctrl", "c")
                        # chrome_app.wait_cpu_usage_lower()  # necessary
                        wait_cpu_usage_lower()
                        new_url = ""
                        new_url = pyperclip.paste()
                        if len(tab_urls) > 0 and (tab_urls[0] == new_url):
                            loop_complete = True  # to be removed
                            break

                        tab_urls.append(new_url)
                        total_tabs+=1
                        print(new_url)
                        pyautogui.hotkey(*BR.tab_switch)
                        # chrome_app.wait_cpu_usage_lower()  # necessary
                        wait_cpu_usage_lower()
                    # chrome_window.urls.append(new_url)
                    all_urls = tab_urls
                    win32gui.SetWindowPlacement(wnd.hwnd, wnd.placement)
        # fm.Close()
        return all_urls

    def GetAllTabUrls(self, windows, ignore_minimized=True):

        # fm=fullscreen_message("Please wait...")
        all_urls = {}
        for wnd in [x for x in windows if GetFilenameFromFullAddress(x.exe) in self.name]:
            # win32gui.ShowWindow()
            # chrome_app = pywinauto.application.Application()
            # chrome_app.connect(handle=wnd.hwnd)
            # chrome_window = chrome_app.windows(handle=wnd.hwnd)[0]
            print(wnd.is_minimized(), ignore_minimized, (not wnd.is_minimized()) or (not ignore_minimized), wnd.title)
            # print(wnd)
            if (not wnd.is_minimized()) or (not ignore_minimized):  # only focus on nonminimized windows. Maybe this should be optional???
                print("--->", wnd)
#
#                 # region BRING WINDOW FOREGROUND
#                 while (win32gui.GetForegroundWindow() != wnd.hwnd):  # makesure the target window is in the foreground
#                     print("fore:", win32gui.GetForegroundWindow(), "wanted:", wnd.hwnd)
#                     win32gui.ShowWindow(wnd.hwnd,win32con.SW_MINIMIZE)
#                     wait_cpu_usage_lower()
#
#                     if (win32gui.GetForegroundWindow() != wnd.hwnd):
#                         win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) #necassary to avoid error...Dont as why, Dont ask how I found it, I just felt it should work.
#                         win32api.SetCursorPos((-10000, 500))
#                         if wnd.is_minimized():
#                             if wnd.was_maximized():
#                                 win32gui.ShowWindow(wnd.hwnd, win32con.SW_MAXIMIZE)
#                             else:
#                                 win32gui.ShowWindow(wnd.hwnd, win32con.SW_RESTORE)
#                         else:
#                             win32gui.ShowWindow(wnd.hwnd, win32con.SW_SHOW)
#                         wait_cpu_usage_lower()
#                         win32gui.SetForegroundWindow(wnd.hwnd)
#                     sleep(1)
#                     wait_cpu_usage_lower()
#                     # win32gui.SetFocus(wnd.hwnd) #access denied?
#                     # wait_cpu_usage_lower()
#                     # chrome_window.minimize()
#                     # chrome_app.wait_cpu_usage_lower()  # necessary
#                     # chrome_window.restore()
#                     # chrome_window.set_focus()
#                     # chrome_app.wait_cpu_usage_lower()  # necessary
#                     # sleep(1) #necessary
#                     # win32gui.ShowWindow(wnd.hwnd,win32con.SW_NORMAL)
#                     # win32gui.BringWindowToTop(wnd.hwnd)
#                     # win32gui.SetForegroundWindow(wnd.hwnd)
#                     tries=0
#                     while (win32gui.GetForegroundWindow() != wnd.hwnd):
#                         # win32gui.ShowWindow(win32gui.GetForegroundWindow(),win32con.SW_FORCEMINIMIZE)
#                         # win32api.SetCursorPos((-10000, 500))
#                         # win32gui.ShowWindow(wnd.hwnd, win32con.SW_SHOW)
#                         # win32gui.BringWindowToTop(wnd.hwnd)
#                         tries+=1
#                         wait_cpu_usage_lower()
#                         if tries>3:
#                             raise Exception("Cannot focus on window " + str(wnd.hwnd))
# # win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE) #minimize this window and try again
#                         # chrome_window.set_focus()
#
#                 # endregion
#                 # chrome_app.wait_cpu_usage_lower()  # necessary
#                 wait_cpu_usage_lower()
#                 # sleep(0.5)
                window.FocusWindow(wnd.hwnd)
                loop_complete = False
                tab_titles = []
                tab_urls = []
                total_trial=0
                while (not loop_complete):
                    total_trial+=1
                    if total_trial==Browser.MAXTABNO:
                        break
                    # tab_titles.append(chrome_window.element_info.name)
                    # win32api.keybd_event()
                    pyautogui.hotkey(*self.bar_select)
                    # chrome_app.wait_cpu_usage_lower()  # necessary
                    pyperclip.copy("")
                    wait_cpu_usage_lower()
                    pyautogui.hotkey("ctrl", "c")
                    #TODO: Take the window title along with the url,
                    # then use it to compare with already open windows when loading sessions,
                    # if a window has same title with titles saved in the seession_window, then
                    # maybe the rest of the tabs are also identical.
                    # Then open that window, cycle to see if there are other websites open too,
                    # if not, when the cycling is complete, open the missing sites.

                    # chrome_app.wait_cpu_usage_lower()  # necessary
                    wait_cpu_usage_lower()
                    new_url=""
                    new_url = pyperclip.paste()
                    if len(tab_urls) > 0 and (tab_urls[0] == new_url):
                        loop_complete = True  # to be removed
                        break
                    if new_url in tab_urls: #necessary for empty or repeating tabs!!! #new_url=="" creates issue with explorer.exe
                        pyautogui.hotkey(*self.tab_switch)
                        wait_cpu_usage_lower()
                        continue


                    tab_urls.append(new_url)
                    print(new_url)
                    pyautogui.hotkey(*self.tab_switch)
                    # chrome_app.wait_cpu_usage_lower()  # necessary
                    wait_cpu_usage_lower()
                # chrome_window.urls.append(new_url)
                all_urls[wnd.hwnd] = tab_urls
                win32gui.SetWindowPlacement(wnd.hwnd,wnd.placement)

                # all_urls.append(tab_urls)
        self.tab_urls=all_urls
        # fm.Close()
        return all_urls



"""--------------------END: BROWSERS-----------------------"""






"""--------------------BEGIN: ALL WINDOWS-----------------------"""



"""
1.Get all window info,
2.Get all com files info
3.if window info is in comfiles regard as com file
4.if window exe is chrome.exe regards as chrome browser
   
"""

class window():

    associated_com_object_info=None
    associated_file_address=None
    tab_urls=[]
    placement=None
    foreground_window=False


    def __init__(self,hwnd,thread_id,pid,title):
        self.hwnd,self.thread_id,self.pid,self.title=hwnd,thread_id,pid,title
        self.exe=self.FindProcessExe()

    def __eq__(self, other):
        if type(other) is window:
            if self.hwnd==other.hwnd and self.title==other.title and self.thread_id==other.thread_id and self.pid==other.pid:
                return True
            return False
        return False



    def was_maximized(self):
        """Indicate whether the window was maximized before minimizing or not"""
        if self.hwnd:
            (flags, _, _, _, _) = win32gui.GetWindowPlacement(self.hwnd)
            return (flags & win32con.WPF_RESTORETOMAXIMIZED == win32con.WPF_RESTORETOMAXIMIZED)
        else:
            return None

    def FindProcessExe(self):
        return psutil.Process(self.pid).exe()

    def is_identical_to(self,other):
        if self.associated_com_object_info == other.associated_com_object_info:
            if self.exe == other.exe:
                if self.associated_file_address==other.associated_file_address:
                    if self.placement==other.placement:
                        if self.tab_urls==other.tab_urls:
                            return True
        return False

    def __repr__(self):

        if self.associated_com_object_info:
            return "COM window - HWND: "+str(self.hwnd)+" - PID: "+str(self.pid)+ " - Title: "+self.title+" - Exe: "+self.exe+" - COM: "+str(self.associated_com_object_info)
        elif self.associated_file_address:
            return "File window - HWND: " + str(self.hwnd) + " - PID: " + str(
                self.pid) + " - Title: " + self.title + " - Exe: " + self.exe + " - File: " + str(
                self.associated_file_address)
        elif self.tab_urls!=[]:
            return "Browser window - HWND: " + str(self.hwnd) + " - PID: " + str(
                self.pid) + " - Title: " + self.title + " - Exe: " + self.exe + " - Tab urls: "+str(self.tab_urls)
        else:
            return "window - HWND: " + str(self.hwnd) + " - PID: " + str(
                self.pid) + " - Title: " + self.title + " - Exe: " + self.exe

    def is_minimized(self):
        return (win32gui.GetWindowPlacement(self.hwnd)[1] == win32con.SW_SHOWMINIMIZED)



    @staticmethod
    def FindAllWindows(only_visible=True,only_titled=True,associate_COMS=True,ignore_minimized=False,ignore_minimized_browsers=True,ignore_urls=False):
        #fixme:ignore minimized do not do anything different
        #TODO: remove identical windows from the list ifonly they are also identical in their window positions as well as associated files etc.



        # USES: win32gui, win32process
        # windows_hwnds = []
        # win32gui.EnumWindows(lambda hwnd, resultList: resultList.append(hwnd), windows_hwnds)
        # windows = [(x, *win32process.GetWindowThreadProcessId(x), win32gui.GetWindowText(x)) for x in windows_hwnds if
        #            (not only_visible or win32gui.IsWindowVisible(x) and (win32gui.GetWindowText(x)!="" or not only_titled))]
        # windows = [window(*x) for x in windows]
        #
        # for wnd in windows:
        #     wnd.placement=win32gui.GetWindowPlacement(wnd.hwnd) #to set it back use: win32gui.SetWindowPlacement(w[0].hwnd,w[0].placement)
        windows=window._FindAllWindows(only_visible=only_visible,only_titled=only_titled,ignore_minimized=ignore_minimized)
        tophwnd=win32gui.GetForegroundWindow()


        #region associate with com items
        if associate_COMS:
            windows=window._AssociateCOMObjects(windows)
            # coms=com_object_info.FindRunningComObjectsAsInfo()
            # for com in coms:
            #     association_size=[(SequenceMatcher(lambda z: z == " ", x.title, com.filename).get_matching_blocks()[0].size, x.title, com.filename,windows.index(x),x) for x in windows if SequenceMatcher(lambda z: z == " ", x.title, com.filename).ratio() > 0.0]
            #     association_size.sort(reverse=True,key=lambda x:x[0])
            #     for assoc in association_size:
            #         if windows[assoc[-2]].associated_com_object_info!=None : #if athis window is already associated
            #             continue
            #         windows[assoc[-2]].associated_com_object_info=com
            #         # print(str(assoc[-1]))
            #         break
        #endregion

        #region associate title file names
        for wnd in windows:
            wnd._FindFileFromTitle()
        #endregion
        #region find chrome tab urls
        if not ignore_urls:
            for browser_name in BROWSERS:

                all_tab_urls=Browser(browser_name).GetAllTabUrls(windows,ignore_minimized_browsers)
                for wnd in windows:
                    try:
                        wnd.tab_urls=all_tab_urls[wnd.hwnd]
                    except:
                        pass

        # for hwnd,val in all_tab_urls.items():
        #     windows[hwnd].tab_urls=val
        #endregion


        if win32gui.GetForegroundWindow()!=tophwnd and win32gui.IsWindowVisible(tophwnd):
            window.FocusWindow(tophwnd)
        for wnd in windows:
            try:
                win32gui.SetWindowPlacement(wnd.hwnd,wnd.placement)
            except:
                print("Error 482 - cannot set window placement")

        to_be_popped=[]
        for i in range(len(windows)):
            for j in range(i):
                if windows[i].is_identical_to(windows[j]):
                    to_be_popped.append(j)
        windows=PopMultipleFromList(windows,to_be_popped)


        return windows

    @staticmethod
    def FindAllWindowsEndingTitleWith(title, only_visible=True):
        windows = window.FindAllWindows(only_visible)
        titled_windows = [x for x in windows if x.title.endswith(title)]
        return titled_windows

    @staticmethod
    def FindAllWindowsContainingInTitle(title, only_visible=True):
        windows = window.FindAllWindows(only_visible)
        titled_windows = [x for x in windows if x.title.__contains__(title)]
        return titled_windows

    def ExtractAssociatedFileAddressFromTitle(self):
        from itertools import combinations
        if self.associated_com_object_info==None:
            combs=combinations(range(len(self.title.split(" "))),2)
            addresses=[" ".join(self.title.split(" ")[x[0]:x[1]]) for x in combs]

            files=[x for x in addresses if os.path.isfile(x)==True]
            if files and len(files)>0:
                files.sort(reverse=True, key=lambda x: len(x))

                self.associated_file_address=GetAbsolutePath(files[0])
                print("HERE:",self.associated_file_address)

    @staticmethod
    def _FindAllWindows(only_visible=True,only_titled=True,ignore_minimized=True):
        # USES: win32gui, win32process
        windows_hwnds = []
        win32gui.EnumWindows(lambda hwnd, resultList: resultList.append(hwnd), windows_hwnds)
        windows = [(x, *win32process.GetWindowThreadProcessId(x), win32gui.GetWindowText(x)) for x in windows_hwnds if
                   (not only_visible or win32gui.IsWindowVisible(x) and (win32gui.GetWindowText(x)!="" or not only_titled))]
        windows = [window(*x) for x in windows ]

        fhwnd=win32gui.GetForegroundWindow()
        for wnd in windows:
            wnd.placement=win32gui.GetWindowPlacement(wnd.hwnd) #to set it back use: win32gui.SetWindowPlacement(w[0].hwnd,w[0].placement)
            if wnd.hwnd==fhwnd:
                wnd.foreground_window=True
        if ignore_minimized:
            wnds=[]
            for wnd in windows:
                if wnd.is_minimized()==False:
                    wnds.append(wnd)
            windows=wnds
        return windows[:-1] #fixme: removing the last element is a quick fix to remove the "explorer.exe" from the list.
                            # Find a better solution later

        #TODO: ignore windows that do not show in screen, because their window width or height is zero or negative.
        # This would eliminates Backup and Sync and Malwarebytes Tray Application for me. Maybe not so necessary.
    

    @staticmethod
    def _AssociateCOMObjects(windows):
        assert type(windows) is list, "_AssociateCOMObjects() requires list of windows"
        assert type(windows[0]) is window, "_AssociateCOMObjects() requires list of windows"
        coms = com_object_info.FindRunningComObjectsAsInfo()
        for com in coms:
            association_size = [(SequenceMatcher(lambda z: z == " ", x.title, com.filename).get_matching_blocks()[
                                     0].size, x.title, com.filename, windows.index(x), x) for x in windows if
                                SequenceMatcher(lambda z: z == " ", x.title, com.filename).ratio() > 0.0]
            association_size.sort(reverse=True, key=lambda x: x[0])
            for assoc in association_size:
                if windows[assoc[-2]].associated_com_object_info != None:  # if athis window is already associated
                    continue
                windows[assoc[-2]].associated_com_object_info = com
                # print(str(assoc[-1]))
                break
        return windows

    def _FindFileFromTitle(wnd):
        assert type(wnd) is window, "_FindFileFromTitle receives window class object as parameter."
        if wnd.associated_com_object_info == None: #if the file is not a com object maybe we can find a associated file from the text
            wnd.ExtractAssociatedFileAddressFromTitle()
        return wnd

    @staticmethod
    def GetForegroundWindow(associate_COMS=True,ignore_minimized_browsers=True):

        #region get all basic info
        # windows_hwnds = []
        # win32gui.EnumWindows(lambda hwnd, resultList: resultList.append(hwnd), windows_hwnds)
        # windows = [(x, *win32process.GetWindowThreadProcessId(x), win32gui.GetWindowText(x)) for x in windows_hwnds if
        #            ( win32gui.IsWindowVisible(x) and (win32gui.GetWindowText(x)!="" ))]
        # windows = [window(*x) for x in windows]
        #endregion

        # for wnd in windows:
        #     wnd.placement=win32gui.GetWindowPlacement(wnd.hwnd) #to set it back use: win32gui.SetWindowPlacement(w[0].hwnd,w[0].placement)
        windows = window._FindAllWindows()

        if associate_COMS:
            windows = window._AssociateCOMObjects(windows)
            # coms=com_object_info.FindRunningComObjectsAsInfo()
            # for com in coms:
            #     association_size=[(SequenceMatcher(lambda z: z == " ", x.title, com.filename).get_matching_blocks()[0].size, x.title, com.filename,windows.index(x),x) for x in windows if SequenceMatcher(lambda z: z == " ", x.title, com.filename).ratio() > 0.0]
            #     association_size.sort(reverse=True,key=lambda x:x[0])
            #     for assoc in association_size:
            #         if windows[assoc[-2]].associated_com_object_info!=None : #if athis window is already associated
            #             continue
            #         windows[assoc[-2]].associated_com_object_info=com
            #         # print(str(assoc[-1]))
            #         break

        #region associate title file names
        for wnd in windows:
            wnd._FindFileFromTitle()
        #endregion

        #look for urls only if its on the foreground
        for wnd in windows:
            if wnd.hwnd==win32gui.GetForegroundWindow():
                wnd.tab_urls=Browser.GetThisTabUrls(wnd)
                return wnd
        return None

    @staticmethod
    def FocusWindow(hwnd):
        try:
            ### BELOW CODE REMOVES FOREGROUND LOCK... UNSURE BUT MAY BE USEFUL TO AVOID ALWAYSONTOP WINDOWS
            win32gui.SystemParametersInfo(win32con.SPI_SETFOREGROUNDLOCKTIMEOUT, 0,
                                          win32con.SPIF_SENDWININICHANGE | win32con.SPIF_UPDATEINIFILE)
            ##########

            while (win32gui.GetForegroundWindow() != hwnd):  # makesure the target window is in the foreground
                print("fore:", win32gui.GetForegroundWindow(), "wanted:", hwnd)
                # win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                wait_cpu_usage_lower()

                if (win32gui.GetForegroundWindow() != hwnd):
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0,
                                         0)  # necassary to avoid error...Dont as why, Dont ask how I found it, I just felt it should work.
                    win32api.SetCursorPos((-10000, 500))
                    if (win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMINIMIZED):
                        if (win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMAXIMIZED):
                            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                        else:
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    else:
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    wait_cpu_usage_lower()
                    win32gui.SetForegroundWindow(hwnd)

                wait_cpu_usage_lower()
                tries = 0
                while (win32gui.GetForegroundWindow() != hwnd):
                    tries += 1
                    wait_cpu_usage_lower()
                    if tries > 3:
                        raise Exception("Cannot focus on window " + str(hwnd))
            # endregion
            wait_cpu_usage_lower()
            return True
        except:
            print("ERROR 639: Cannot focus window ",win32gui.GetWindowText(hwnd))
            return False



"""--------------------END: ALL WINDOWS-----------------------"""




"""--------------------BEGIN: FULLSCREEN MESSAGE-----------------------"""



FULLSCREEN_MESSAGE = ""
FULLSCREEN_TEXTCOLOR = (0, 0, 0)

class fullscreen_message():
    className=""
    hwnd=0
    wndClass=None
    wndClassAtom=None
    hInstance = None
    Alpha=200
    def __init__(self,message,textcolor=(150, 0, 55),alpha=200,auto_start=False,window_class_name="FullscreenMessage"):
        # self.FULLSCREEN_MESSAGE,self.FULLSCREEN_TEXTCOLOR=message,textcolor
        global FULLSCREEN_TEXTCOLOR,FULLSCREEN_MESSAGE
        FULLSCREEN_TEXTCOLOR, FULLSCREEN_MESSAGE= textcolor, message
        self.Alpha=alpha
        self.className=window_class_name+RemovePuctuations(message)
        self.hInstance = win32api.GetModuleHandle()
        # self.hInstance = 71

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633576(v=vs.85).aspx
        # win32gui does not support WNDCLASSEX.
        self.wndClass                = win32gui.WNDCLASS()
        # http://msdn.microsoft.com/en-us/library/windows/desktop/ff729176(v=vs.85).aspx
        self.wndClass.style          = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        self.wndClass.lpfnWndProc    = self.wndProc
        self.wndClass.hInstance      = self.hInstance
        self.wndClass.hCursor        = win32gui.LoadCursor(None, win32con.IDC_ARROW)
        self.wndClass.hbrBackground  = win32gui.GetStockObject(win32con.WHITE_BRUSH)
        self.wndClass.lpszClassName  = self.className
        # win32gui does not support RegisterClassEx
        self.wndClassAtom = win32gui.RegisterClass(self.wndClass)


        if auto_start:
            self.Show()

    def Show(self):

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ff700543(v=vs.85).aspx
        # Consider using: WS_EX_COMPOSITED, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT
        # The WS_EX_TRANSPARENT flag makes events (like mouse clicks) fall through the window.
        exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632600(v=vs.85).aspx
        # Consider using: WS_DISABLED, WS_POPUP, WS_VISIBLE
        style = win32con.WS_DISABLED | win32con.WS_POPUP | win32con.WS_VISIBLE

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms632680(v=vs.85).aspx
        hWindow = win32gui.CreateWindowEx(
            exStyle,
            self.wndClassAtom,
            None, # WindowName
            style,
            0, # x
            0, # y
            win32api.GetSystemMetrics(win32con.SM_CXSCREEN), # width
            win32api.GetSystemMetrics(win32con.SM_CYSCREEN), # height
            None, # hWndParent
            None, # hMenu
            self.hInstance,
            None # lpParam
        )

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633540(v=vs.85).aspx
        win32gui.SetLayeredWindowAttributes(hWindow, 0x00ffffff, self.Alpha, win32con.LWA_COLORKEY | win32con.LWA_ALPHA)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/dd145167(v=vs.85).aspx
        win32gui.UpdateWindow(hWindow)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633545(v=vs.85).aspx
        win32gui.SetWindowPos(hWindow, win32con.HWND_TOPMOST, 0, 0, 0, 0,
            win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

        # http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
        win32gui.ShowWindow(hWindow, win32con.SW_SHOW)
        # win32gui.PumpMessages()
        self.hwnd=hWindow
        return hWindow

    def Close(self):

        win32gui.CloseWindow(self.hwnd)
        win32gui.DestroyWindow(self.hwnd)
        win32gui.UnregisterClass(self.wndClassAtom,self.hInstance)

    def Hide(self):
        print("fm:",self.hwnd)
        win32gui.ShowWindow(self.hwnd,win32con.SW_HIDE)


    @staticmethod
    def wndProc(hWnd, message, wParam, lParam):
        if message == win32con.WM_PAINT:
            hdc, paintStruct = win32gui.BeginPaint(hWnd)

            dpiScale = win32ui.GetDeviceCaps(hdc, win32con.LOGPIXELSX) / 60.0
            fontSize = 80

            # http://msdn.microsoft.com/en-us/library/windows/desktop/dd145037(v=vs.85).aspx
            lf = win32gui.LOGFONT()
            lf.lfFaceName = "Arial"
            lf.lfHeight = int(round(dpiScale * fontSize))

            global FULLSCREEN_TEXTCOLOR
            # FULLSCREEN_TEXTCOLOR=(100,200,0)
            #set color below:
            win32gui.SetTextColor(hdc, win32api.RGB(*FULLSCREEN_TEXTCOLOR))

            # Use nonantialiased to remove the white edges around the text.
            lf.lfQuality = win32con.NONANTIALIASED_QUALITY
            hf = win32gui.CreateFontIndirect(lf)

            win32gui.SelectObject(hdc, hf)

            rect = win32gui.GetClientRect(hWnd)
            # http://msdn.microsoft.com/en-us/library/windows/desktop/dd162498(v=vs.85).aspx

            global FULLSCREEN_MESSAGE
            # FULLSCREEN_MESSAGE="Please wait..."
            win32gui.DrawText(
                hdc,
                FULLSCREEN_MESSAGE,
                -1,
                rect,
                # win32con.DT_CENTER | win32con.DT_NOCLIP | win32con.DT_SINGLELINE | win32con.DT_VCENTER
                win32con.DT_CENTER | win32con.DT_NOCLIP | win32con.DT_SINGLELINE | win32con.DT_VCENTER
            )
            win32gui.EndPaint(hWnd, paintStruct)
            return 0

        elif message == win32con.WM_DESTROY:
            print ('Closing the window.')
            win32gui.PostQuitMessage(0)
            return 0

        else:
            return win32gui.DefWindowProc(hWnd, message, wParam, lParam)

    # def ShowFullScreenMessage():
    #     win32gui.PumpMessages()

"""--------------------END: FULLSCREEN MESSAGE-----------------------"""


"""--------------------BEGIN: DESKTOP SESSIONS-----------------------"""

class desktop_session():





    def __init__(self,windows,pass_hwnd=False):

        self.session_windows=[self.session_window(w,pass_hwnd=pass_hwnd) for w in windows]
        n=datetime.datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
        self.session_date=n
        self.session_priority=3
        self.deadline=""
        self.reminders=True


    def __iter__(self):
        self.n=0
        return self

    def __next__(self):
        if self.n < self.session_windows.__len__():
            result=self.session_windows[self.n]
            self.n += 1
            return result
        else:
            raise StopIteration

    def Update(self,windows):
        self.session_windows = [self.session_window(w) for w in windows]

    def AddTopWindow(self):
        hwnd=win32gui.GetForegroundWindow()
        wnd=window(hwnd,*win32process.GetWindowThreadProcessId(hwnd), win32gui.GetWindowText(hwnd))
        sw=self.session_window(wnd)

        self.session_windows.append(sw)

    def __add__(self, other):
        if type(other) is desktop_session.session_window:
            desktop_session.session_windows.append(other)

    def __contains__(self, item):
        if type(item) is desktop_session.session_window:
            for w in self.session_windows:
                if w==item:
                    return True
            return False
        elif type(item) is window:
            item=desktop_session.session_window(item)
            for w in self.session_windows:
                if w==item:
                    return True
            return False
        Warning("type not suitable for comparison.")
        return False

    def DumpToFile(self,filename):
        pickle.dump(self, open(filename, "wb"))

    @staticmethod
    def LoadFromFile(filename):
        try:
            return pickle.load(open(filename,"rb"))
        except:
            return

    class session_window():
        swTopWindow=False
        swPlacement=None
        swExe=None
        swBrowserUrls=[]
        swComData=None
        swAssociatedFileAddress=None
        def __init__(self,wnd:window,pass_hwnd=False):
            self.swPlacement=wnd.placement
            self.swExe=wnd.exe
            self.swAssociatedFileAddress=wnd.associated_file_address
            if wnd.associated_com_object_info!=None:
                self.swComData=wnd.associated_com_object_info.full_address
            # self.swComData = wnd.associated_com_object_info
            self.swBrowserUrls=wnd.tab_urls
            if pass_hwnd:
                self.hwnd=wnd.hwnd
            self.swTopWindow=wnd.foreground_window

        def __eq__(self, other):
            assert type(self)==type(other), "different types. Cannot be checked for equality!"
            if self.swExe==other.swExe and set(self.swBrowserUrls)==set(other.swBrowserUrls) and self.swComData==other.swComData and self.swAssociatedFileAddress==other.swAssociatedFileAddress:
                return True
            else:
                return False

        def identical(self,other): #considers placement As Well
            assert type(self)==type(other), "different types. Cannot be checked for equality!"
            if self.swExe==other.swExe and set(self.swBrowserUrls)==set(other.swBrowserUrls) and self.swComData==other.swComData and self.swAssociatedFileAddress==other.swAssociatedFileAddress and self.swPlacement==other.swPlacement:
                return True
            else:
                return False

        def DeleteWindowFromMenu(self,systray):
            print("-----------")
            for s in systray.sessions:
                for w_idx in range(len(s.session_windows)):

                    if s.session_windows[w_idx] is self:
                        print("found",s.session_windows[w_idx],self )
                        print(len(s.session_windows))
                        s.session_windows.pop(w_idx)
                        print(len(s.session_windows))
                        systray._Update(systray)
                        systray.UpdateMenuOptionsFromTray(systray)
                        systray.SaveSessionsToFileFromMenu(systray)
                        return

        def LoadSessionWindowFromMenu(swnd,systray):
            systray.loading_message.Show()
            swnd.LoadSessionWindow()
            systray.loading_message.Hide()





        def LoadSessionWindow(swnd):


            file = swnd.swComData
            if file == None:
                file = swnd.swAssociatedFileAddress
            if file == None:
                file = ""

            w0 = window._FindAllWindows()
            [print(x) for x in w0]
            print('"' + swnd.swExe + '" "' + file + '"')
            try:
                p = win32process.CreateProcess(None, '"' + swnd.swExe + '" "' + file + '"', None, None, 0,
                                               win32process.DETACHED_PROCESS, None, None, win32process.STARTUPINFO())
                wait_cpu_usage_lower(initial_sleep=0.1) #initial sleep is crucial so that new window is detected by the "_FindAllWindows" in the definition of w1
            except:
                print("Error 960: Cannot start a process of: " + swnd.swExe +" due to process creating error")
                return None
            print("------------------------------")
            w1 = window._FindAllWindows()
            [print(x) for x in w1]
            print("------------------------------")

            newwnd = [i for i in w1 if i not in w0]  # this must be the new window handle
            # print(newwnd)
            # fixme: if the length of newwnd is more than 1 the other window might be associated with another window in the list!

            try:
                newwnd = newwnd[0]
            except:
                print(newwnd)
                print("Error 972: Cannot start a process of: " + swnd.swExe + " since unable to find the HWND")
                return None
            # if newwnd in finalized_hwnds: print("<<<>>>this window is repeated:", win32gui.GetWindowText(newwnd))
            # finalized_hwnds.append(newwnd)

            win32gui.SetWindowPlacement(newwnd.hwnd, swnd.swPlacement)

            wait_cpu_usage_lower()

            if swnd.swBrowserUrls != []:
                ### THIS PIECE IS FOUND ONLINE TO ASSURE THE WINDOW IS FICUSSED NO MATTER WHAT BUT IT DOESNT WORK
                # win32gui.SetWindowPos(newwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                #                       win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
                # win32gui.SetWindowPos(newwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                #                       win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
                # win32gui.SetWindowPos(newwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                #                       win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
                ##########################################################################################

                first_page = True
                for url in swnd.swBrowserUrls:

                    if not first_page:
                        pyautogui.hotkey("ctrl", "t")
                    wait_cpu_usage_lower()
                    if win32gui.GetForegroundWindow() != newwnd.hwnd:  ### Asserts the window is still top
                        window.FocusWindow(newwnd.hwnd)
                    pyautogui.hotkey("alt", "d")
                    pyperclip.copy(url)
                    wait_cpu_usage_lower()
                    if win32gui.GetForegroundWindow() != newwnd.hwnd:  ### Asserts the window is still top
                        window.FocusWindow(newwnd.hwnd)
                    pyautogui.hotkey("ctrl", "v")
                    wait_cpu_usage_lower()
                    if win32gui.GetForegroundWindow() != newwnd.hwnd:  ### Asserts the window is still top
                        window.FocusWindow(newwnd.hwnd)
                    pyautogui.press("enter")
                    wait_cpu_usage_lower()
                    first_page = False
            # fixme: explorer.exe window had youtube url which should have been under firefox.

            # fixme: Noted an issue where an additional explorer.exe opens up with
            # a title of "Program Manager". it is always the last one.
            # a quick fix is applied, but might not work in the future. Check end of _FindAllWindows
            return newwnd.hwnd


    def DeleteSessionFromMenu(self,systray):

        cnfm=pyautogui.confirm("Are you sure you want to delete the session "+self.session_date,"",["Yes","No"])
        if cnfm=="No":
            return


        for s in systray.sessions:
            if s.session_date==self.session_date:
                systray.sessions.pop(systray.sessions.index(s))
                del s
                systray._Update(systray)
                systray.UpdateMenuOptionsFromTray(systray)
                systray.SaveSessionsToFileFromMenu(systray)
                return



    def RenameSessionFromMenu(self,systray):
        newname=pyautogui.prompt("Enter a new name","Rename "+self.session_date,default=self.session_date)
        if newname==None:
            return

        for s in systray.sessions:
            if s.session_date==self.session_date:
                s.session_date=newname
                # systray.sessions.pop(systray.sessions.index(s))
                # del s
                systray._Update(systray)
                systray.UpdateMenuOptionsFromTray(systray)
                systray.SaveSessionsToFileFromMenu(systray)
                return


    def LoadSessionFromMenu(self,systray):
        if wait_cpu_usage_lower(timeout=1) == "timeout":
            cnfm = pyautogui.alert(
                "The CPU load is too high to load a new session, "+self.session_date+". Please wait until active processes are complete.",
                "Cannot load session",
                ["OK"])
            return

        systray.loading_message.Show()
        print("..",systray.last_activated_session,"..",self.session_date)
        systray.last_activated_session=self.session_date
        self.LoadSession(close_unneeded_windows=CLOSE_UNWANTED)
        systray.loading_message.Hide()
        systray.UpdateMenuOptionsFromTray(systray)
        # fm.Close()




    def LoadSession(self,close_unneeded_windows=False):
        pass
        #Decide which apps to open
        #   if a window is not identical with an existing one reopen.
        #   all existing Browsers are minimized, all session browsers are reopenned.
        #
        e = window.FindAllWindows(ignore_minimized=False, ignore_urls=True)
        ep=desktop_session(e,pass_hwnd=True)

        # fm=fullscreen_message("Loading session...")

        unneeded_windows=[i for i in ep if i not in self]
        if close_unneeded_windows:
            for wnd in unneeded_windows:
                print(wnd.hwnd)
                win32gui.PostMessage(wnd.hwnd,win32con.WM_CLOSE,0,0)

        else:
            for wnd in unneeded_windows:
                try:

                    win32gui.ShowWindow(wnd.hwnd,win32con.SW_MINIMIZE)
                except:
                    print("ERROR 984: Cannot minimize unneeded window ",wnd.swExe)
                    pass

        common_session_windows = [(j, k) for k in ep for j in self if k == j]
        for swnd,wnd in common_session_windows: #NEED TO GET HWND FROM ep
            try:
                win32gui.SetWindowPlacement(wnd.hwnd,swnd.swPlacement) #NOT WORKING
            except:
                try:
                    win32gui.ShowWindow(wnd.hwnd,win32con.SW_MINIMIZE)
                    print("ERROR 994: Cannot set window placement for ", swnd.swExe, wnd.hwnd, "level 1")
                except:
                    print("ERROR 996: Cannot set window placement for ",swnd.swExe,wnd.hwnd,"level 2")
            # if j.swTopWindow:
            #     window.
            #     win32gui.SetForegroundWindow(k.hwnd)

        needed_session_windows = [i for i in self if i not in ep]
        finalized_hwnds=[]


        for swnd in needed_session_windows:
            new_hwnd= swnd.LoadSessionWindow()
            if new_hwnd: finalized_hwnds.append(new_hwnd)



        # if self.parent_menu!=None:
        #     # self.parent_menu.active_session_date=self.session_date
        #     self.parent_menu.PostLoadFunction(self)

        #look for current open windows: Only visible ones.
        #Browsers should not be bothered for checiking identity.!! too long.

        #if any is equal, them adopt its placement



"""--------------------END: DESKTOP SESSIONS-----------------------"""

"""--------------------BEGIN: MENU-----------------------"""

import infi.systray #from https://github.com/MagTun/infi.systray




class menu():


    # SAVEFILE="DesktopSessionsProfile.dsp"
    #TODO: menu class is such a mess. needs to be fixed...but it works

    #TODO: add manual entry

    def __init__(self):
        print("stating desktop sessions manager")
        global PROGRAM_ICO
        global PROGRAM_TITLE
        self.systray=infi.systray.SysTrayIcon(PROGRAM_ICO, PROGRAM_TITLE, (), on_quit=menu.Quit)
        self.systray.SAVEFILE=os.getenv("LOCALAPPDATA")+"\\Avokado\\DesktopSessionsManager\\DesktopSessionsProfile.dsp"

        # self.systray.ignore_minimized = True

        # self.systray.core_tuple = (("Save as new session", None, self.SaveAsNewSession), ("Settings...",None, menu.dummy_func) )
        self.systray.core_tuple = ( ("Save as new session", None, self.SaveAsNewSession),)

        self.systray.last_activated_session=-1
        self.systray.saving_message = fullscreen_message("please wait while saving", auto_start=False) #fixme: if the two fullscreen messages have same message it doesnt work. No mather the firstone is, always the latter one works. I coudnt find a way to fix it.
        self.systray.loading_message= fullscreen_message("please wait...",auto_start=False)
        self.systray.menu_options=()
        self.systray.UpdateMenuOptionsFromTray=self.UpdateMenuOptionsFromTray
        self.systray._Update=self._UpdateFromTray
        self.systray.SaveSessionsToFileFromMenu=self.SaveSessionsToFileFromMenu
        self.LoadSessionsFromFile()
        self.systray.UpdateMenuOptionsFromTray(self.systray)
        self.systray.start()
        # self.systray.CHECK_ICO=self.CHECK_ICO








    @staticmethod
    def dummy_func(self):
        return

    @staticmethod
    def switchSetting_ignore_minimized(systray):
        # systray.ignore_minimized=not systray.ignore_minimized
        global IGNORE_MINIMIZED
        IGNORE_MINIMIZED=not IGNORE_MINIMIZED
        systray.UpdateMenuOptionsFromTray(systray)

    @staticmethod
    def switchSetting_close_unwanted(systray):
        global CLOSE_UNWANTED
        print("ear",CLOSE_UNWANTED)
        CLOSE_UNWANTED=not CLOSE_UNWANTED
        print("aft",CLOSE_UNWANTED)
        systray.UpdateMenuOptionsFromTray(systray)
        systray.SaveSessionsToFileFromMenu(systray)


    @staticmethod
    def ResetMaxCPUloadPercent(systray):
        global MAX_CPU_LOAD
        newval=pyautogui.prompt("Enter the maximum CPU load percent to load or save...",default=MAX_CPU_LOAD)
        try:
            newval=float(newval)
            MAX_CPU_LOAD=newval
            systray.SaveSessionsToFileFromMenu(systray)
        except:
            pass

    @staticmethod
    def DeleteAllSessions(systray):
        a=pyautogui.confirm("Are you sure?","Delete all saved sessions",buttons=["Yes","No"])
        if a=="Yes":
            systray.sessions=[]
            systray.SaveSessionsToFileFromMenu(systray)
            systray.UpdateMenuOptionsFromTray(systray)


    # def UpdateMenuOptions(self):
    #     print(".>>>UpdateMenuOptions",self)
    #
    #     session_menu=[]
    #     for s in self.systray.sessions:
    #
    #         session_option=(s.session_date,self.CHECK_ICO if (s.session_date==self.systray.last_activated_session) else None,s.LoadSessionFromMenu)
    #         session_menu.append( session_option )
    #         session_del_option=("└» delete",self.CHECK_ICO if (s.session_date!=self.systray.last_activated_session) else None,s.DeleteSessionFromMenu)
    #         session_menu.append(session_del_option)
    #         session_menu.append(("-----", None, menu.dummy_func))
    #     self.systray.menu_options=tuple(session_menu)+self.systray.core_tuple
    #     self._Update()

    @staticmethod
    def UpdateMenuOptionsFromTray(systray):
        # try:
        #     systray.loading_message.Hide()
        #     systray.saving_message.Hide()
        #
        # except:
        #     pass

        print(".>>>UpdateMenuOptionsFromTray",systray)
        session_menu=[]
        global CHECK_ICO
        global IGNORE_MINIMIZED

        for s in systray.sessions:
            print("--->>",s.session_date, "",systray.last_activated_session)
            session_option=(s.session_date,CHECK_ICO if (s.session_date==systray.last_activated_session) else None,s.LoadSessionFromMenu)
            session_menu.append( session_option )
            session_rename_option=(" » rename",None,s.RenameSessionFromMenu)
            session_menu.append(session_rename_option)
            session_del_option=(" » delete",None,s.DeleteSessionFromMenu)
            session_menu.append(session_del_option)

            # session_priority_option=(" » priority » ",None,tuple([(str(x),(None if x!=3 else CHECK_ICO),lambda x:x) for x in range(5,0,-1)]+[("turn reminders off",None,lambda x:x),("set deadline...",None,lambda x:x)]))
            # session_menu.append(session_priority_option)

            # session_deadline_option=(" » set deadline date...",None,lambda x:x)
            # session_menu.append(session_deadline_option)

            ###CODE BELOW ADDS A SUBMENU CONTAINING ALL WINDOWS OF THAT SESSION
            # def _delete_session_window(s,)

            sub_session_menu=[]
            for w in s.session_windows:

                #### I AM SURE A BETTER CODE CAN BE WRITTEN!!! TOO MANY TRIALS HERE!!
                try:
                    associated_file_or_link=" - "+GetFilenameFromFullAddress(w.swComData)
                except:
                    try:
                        associated_file_or_link = " - "+w.swBrowserUrls[0]
                    except:
                        # associated_file_or_link= GetFilenameFromFullAddress(w.swAssociatedFileAddress,ignore_nonstring=True)
                        try:
                            associated_file_or_link= " - "+w.swAssociatedFileAddress
                        except:
                            associated_file_or_link=""
                #TODO: complete "load window" function
                sub_session_menu.append((GetFilenameFromFullAddress(w.swExe)+associated_file_or_link ,None, (("load window",None,w.LoadSessionWindowFromMenu), ( "delete window", None, w.DeleteWindowFromMenu), )  ))
            sub_session_menu=( " » all windows » ",None,tuple(sub_session_menu))
            session_menu.append(sub_session_menu)
            #####################

            session_menu.append(("-----", None, menu.dummy_func))


        # systray.settings_tuple= ( ("Settings » ",None, ( ("include minimized during saving", None if systray.ignore_minimized else systray.CHECK_ICO ,menu.switchSetting_ignore_minimized) , ("close other windows before loading a session",None,menu.dummy_func),("always check for open browser tabs instead of opening a new browser",None,menu.dummy_func),("max CPU usage during url saving and loading...",None,menu.dummy_func),("Delete all sessions",None,menu.dummy_func))) ,)
        systray.settings_tuple = (("Settings » ", None, (("include minimized during saving",
                                                          None if IGNORE_MINIMIZED else CHECK_ICO,
                                                          menu.switchSetting_ignore_minimized), (
                                                         "close other windows before loading a session",
                                                         None if not CLOSE_UNWANTED else CHECK_ICO,
                                                         menu.switchSetting_close_unwanted), (
                                                         "max CPU usage during url saving and loading...", None,
                                                         menu.ResetMaxCPUloadPercent),
                                                         ("-----",None,menu.dummy_func)
                                                         ,
                                                         ("Delete all sessions", None, menu.DeleteAllSessions))),)

        systray.menu_options=tuple(session_menu)+systray.core_tuple+systray.settings_tuple
        systray._Update(systray)

    # @staticmethod
    def PostLoadMethod(self,d):
        print(".>>>PostLoadMethod",type(self),type(d))
        # self.last_activated_session=d.session_date
        # self.systray.UpdateMenuOptionsFromTray(systray)
        # self._UpdateFromTray(self)


    # @staticmethod
    # def RenameSession(self,session_date):
    #     pass

    def _Update(self):
        print(".>>>_Update:",self)
        self.systray.update(menu_options=self.systray.menu_options)

    @staticmethod
    def _UpdateFromTray(systray):
        print(".>>>_UpdateFromTray",systray)
        systray.update(menu_options=systray.menu_options)


    @staticmethod
    def UpdateCurrentSession(self):
        pass #TODO:here

    # @staticmethod
    def LoadSessionsFromFile(self):

        try:
            global MAX_CPU_LOAD
            global IGNORE_MINIMIZED
            global CLOSE_UNWANTED
            self.systray.sessions,MAX_CPU_LOAD,IGNORE_MINIMIZED,CLOSE_UNWANTED = pickle.load(open(self.systray.SAVEFILE,"rb"))

        except:
            print("ERROR 1231: No sessions file found")
            self.systray.sessions= []


    # def SaveSessionsToFile(self):
    #     pickle.dump(self.systray.sessions,open(self.SAVEFILE,"wb"))


    def SaveSessionsToFileFromMenu(self,systray):
        if not os.path.isdir(os.getenv("LOCALAPPDATA")+"\\Avokado\\DesktopSessionsManager") :
            os.makedirs(os.getenv("LOCALAPPDATA")+"\\Avokado\\DesktopSessionsManager")
        global MAX_CPU_LOAD
        global IGNORE_MINIMIZED
        global CLOSE_UNWANTED
        pickle.dump((self.systray.sessions,MAX_CPU_LOAD,IGNORE_MINIMIZED,CLOSE_UNWANTED ), open(self.systray.SAVEFILE, "wb"))

    def SaveAsNewSession(self,systray):

        if wait_cpu_usage_lower(timeout=1)=="timeout":

            cnfm = pyautogui.alert ("The CPU load is too high to save a new session. Please wait until active processes are complete.", "Cannot save as new session",
                                     ["OK"])
            return



        # WindowsBalloonTip("Desktop Manager:","Please wait...")
        self.systray.saving_message.Show()
        print("saveassession:",type(self))
        global IGNORE_MINIMIZED
        w=window.FindAllWindows(ignore_minimized=IGNORE_MINIMIZED ,ignore_minimized_browsers=IGNORE_MINIMIZED ) #destoys the tray icon somehow
        d=desktop_session(w)

        self.systray.sessions.append(d)
        self.systray.last_activated_session=d.session_date
        self.systray.saving_message.Hide()
        self.systray.UpdateMenuOptionsFromTray(self.systray)
        self.systray.SaveSessionsToFileFromMenu(self.systray)
        # self.SaveSessionsToFile()

        # menu._Update(self)
        # WindowsBalloonTip("Desktop Manager:", "Please wait...")

    # def AddManualEntry(self,systray):
    #     for s in self.systray.sessions:
    #         if s.session_date==self.systray.last_selection:
    #             new_cmd=pyautogui.prompt("Please enter full command line command for the new window for this session","",["OK","Cancel"])
    #             if new
    #             #TODO: SHOW DIALOG FOR COMMAND LINE ENTRY
    #               # How to distinguish between a word command and pycharm command for instance?
    

    @staticmethod
    def Quit(self):
        sys.exit()



"""--------------------END: MENU-----------------------"""

"""--------------------BEGIN: NOTIFICATION-------------------"""



class WindowsBalloonTip:
    def __init__(self):

        message_map = {
            win32con.WM_DESTROY: self.Destroy,
        }
        # Register the Window class.
        wc = win32gui.WNDCLASS()
        self.hinst = wc.hInstance = win32gui.GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map  # could also specify a wndproc.
        self.classAtom = win32gui.RegisterClass(wc)
        # Create the Window.
    def Show(self,title, msg):
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(self.classAtom, "Taskbar", style, \
                                 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                                 0, 0, self.hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join(sys.path[0], "balloontip.ico"))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            hicon = win32gui.LoadImage(self.hinst, iconPathName, \
                              win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER + 20, hicon, "tooltip")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, \
                         (self.hwnd, 0, win32gui.NIF_INFO, win32con.WM_USER + 20, \
                          hicon, "Balloon  tooltip", msg, 200, title))
        # self.show_balloon(title, msg)
        # time.sleep(10)
        # win32gui.DestroyWindow(self.hwnd)
    def Destroy(self):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

    # def OnDestroy(self, hwnd, msg, wparam, lparam):
    #     nid = (self.hwnd, 0)
    #     win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
    #     win32gui.PostQuitMessage(0)  # Terminate the app.

"""-------------------END: NOTIFICATION--------------------"""



if __name__=="__main__":

    m=menu()

    # w=window.FindAllWindows(ignore_minimized_browsers=False)

    # print(w)

    #TODO: send reminder as notification according to priority
    #TODO: cpu threshold should consider core count
    #TODO: Manage a menu hierarchy (optional)
    #TODO: otherfucntions to maintain profiles
    # ->Session name
    # ->options->
    #   ->Rename...
    #   ->Add the top window
    #   ->Manual entry...
    #   ->Delete (confirmation required once)
    #      ->Priority=med ->
    #         ->high
    #         ->med
    #         ->low
    #         ->none
    #   ->Update (confirmation required)


    #TODO: closeall button



"""--------------------BEGIN: USAGE----------------------

Call window.FindAllWindows()
It will define a list of windows with automatically associate to COM objects pr files in the title



--------------------END: USAGE----------------------"""

#TODO: Generate a list of tests!!!
#todo: tests with two windowed applicaitons


def __unittest__1():
    w=window.FindAllWindows(ignore_minimized=False)
    print("Please perform changes for 5 sec.")
    sleep(5)
    e = window.FindAllWindows(ignore_minimized=False, ignore_urls=True)
    d = desktop_session(w)
    ed = desktop_session(e,pass_hwnd=True)
    com = [(j, k) for k in ed for j in d if k == j]
    for j, k in com:  # NEED TO GET HWND FROM ep
        win32gui.SetWindowPlacement(k.hwnd, k.swPlacement)  # NO CHANGE
    print("No reversion must be seen yet")
    input("Enter to continue")
    for j, k in com:  # NEED TO GET HWND FROM ep
        win32gui.SetWindowPlacement(k.hwnd, j.swPlacement)  # CHANGE
    print("The changes must have been reverted")
    print("test complete.")

def __unittest__2():

    sleep(2)
    print("stating test")
    w = window.FindAllWindows()
    d = desktop_session(w)
    print("start test changes..",end="")
    sleep(1)
    print("1..",end="")
    sleep(1)
    print("2..", end="")
    sleep(1)
    print("3..", end="")
    sleep(1)
    print("4..", end="")
    sleep(1)
    print("5")

    d.LoadSession()
    print("test complete.")