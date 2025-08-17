# 导入库，相当于买了机械手的零件
import pyautogui
import time
import win32api
import win32con
import win32gui
import win32clipboard as w
import pyperclip  # 导入pyperclip库
from wxauto import *
from uiautomation import WindowControl

# 制作机械手 
def FindWindow(chatroom):
    win = win32gui.FindWindow('WeChatMainWndForPC', chatroom)
    print("找到窗口句柄：%x" % win)
    if win != 0:
        win32gui.ShowWindow(win, win32con.SW_SHOWMINIMIZED)
        win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
        win32gui.ShowWindow(win, win32con.SW_SHOW)
        win32gui.SetWindowPos(win, win32con.HWND_TOP, 0, 0, 500, 700, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(win)  # 获取控制
        time.sleep(1)
        tit = win32gui.GetWindowText(win)
        print('已启动【' + str(tit) + '】窗口')
    else:
        print('找不到【%s】窗口' % chatroom)
        exit()
 
# 设置和粘贴剪贴板
def ClipboardText(ClipboardText):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, ClipboardText)
    w.CloseClipboard()
    time.sleep(1)
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(86, 0, 0, 0)
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
 
# 模拟发送动作
def SendMsg():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
 
def SendWxMsg(wxid, sendtext, xunhuan):
    # 先启动微信
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Msg = WindowControl(Name='微信')
    wx_Msg.ButtonControl(Name='聊天').Click()
    # 定位到搜索框
    pyautogui.moveTo(143, 39)
    pyautogui.click()
    # 搜索微信
    ClipboardText(wxid)
    time.sleep(1)
    # 进入聊天窗口
    pyautogui.moveTo(155, 120)
    pyautogui.click()
    # 粘贴文本内容 发送
    for i in range(xunhuan):
        ClipboardText(sendtext)
        SendMsg()
        time.sleep(0.01)
    print('已给'+wxid+'发送'+sendtext)

def SendWxFile(wxid,filepath,xunhuan):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_File = WindowControl(Name='微信')
    wx_File.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(179,53)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(150,149)
    pyautogui.click()
    # 输入路径
    file_path = filepath
    # 选择群聊
    for i in range(xunhuan):
      # 点击发送文件
      wx_File.ButtonControl(Name='发送文件').Click()
      # 发送文件
      # 输入文件地址
      pyperclip.copy(file_path)  # 使用pyperclip将文件路径复制到剪贴板
      pyautogui.hotkey('ctrl', 'v')  # 使用快捷键粘贴文件路径
      time.sleep(1)
      pyautogui.press('enter')
      # 发送
      pyautogui.press('enter')
    print('已给'+wxid+'发送'+filepath)

def WxVoiceCall(wxid,mode):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Voice = WindowControl(Name='微信')
    wx_Voice.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(179,53)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(150,149)
    pyautogui.click()
    wx_Voice.ButtonControl(Name=mode).Click()
    print('已给'+wxid+'拨打了'+mode)
    Guaduan=input(',需要挂断吗？(Y/N):')
    if Guaduan == 'Y':
        pyautogui.moveTo(951,820)
        pyautogui.click()

def SearchWxid(wxid):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Search = WindowControl(Name='微信')
    wx_Search.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(179,53)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(150,149)
    pyautogui.click()
    print('已找到'+wxid)

def AddWx(wxid):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Add = WindowControl(Name='微信')
    wx_Add.ButtonControl(Name='通讯录').Click()
    wx_Add.ButtonControl(Name='添加朋友').Click()
    pyautogui.moveTo(143,53)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.moveTo(219,126)
    pyautogui.click()
    pyautogui.moveTo(475,244)
    pyautogui.click()
    pyautogui.moveTo(462,254)
    pyautogui.click()
    print('已给'+wxid+'发送申请')

def DelWx(wxid):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Del = WindowControl(Name='微信')
    wx_Del.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(217,52)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(210,143)
    pyautogui.click()
    wx_Del.ButtonControl(Name='聊天信息').Click()
    pyautogui.moveTo(929,64)
    pyautogui.click()
    wx_Del.ButtonControl(Name='更多').Click()
    pyautogui.moveTo(1305,298)
    pyautogui.click()
    pyautogui.moveTo(1068,363)
    pyautogui.click()
    print('已删除'+wxid)

def SearchWxHistory(wxid,type):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_History1 = WindowControl(Name='微信')
    wx_History1.ButtonControl(Name='聊天').Click()
    wx_History = WindowControl(Name='微信')
    wx_History.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(179,53)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(150,149)
    pyautogui.click()
    wx_History.ButtonControl(Name='聊天记录').Click()
    if type == 'word':
        pyautogui.moveTo(663,221)
        pyautogui.click()
    elif type == 'photo_and_video':
        pyautogui.moveTo(752,217)
        pyautogui.click()
    elif type == 'web_page':
        pyautogui.moveTo(832,224)
        pyautogui.click()
    elif type == 'music':
        pyautogui.moveTo(905,225)
        pyautogui.click()
    elif type == 'program':
        pyautogui.moveTo(1017,223)
        pyautogui.click()
    elif type == 'video_id':
        pyautogui.moveTo(1096,217)
        pyautogui.click()
    elif type == 'date':
        pyautogui.moveTo(1156,219)
        pyautogui.click()
    else:
        print('错误:未找到分类')
    print('已找到'+wxid+'的'+type+'类型的聊天记录')

def GetMeWx():
    FindWindow('微信')
    time.sleep(1)
    pyautogui.moveTo(35,61)
    pyautogui.click()
    print('已找到自己的微信信息')

def GetFriendWx(wxid):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_GetFriend = WindowControl(Name='微信')
    wx_GetFriend.ButtonControl(Name='通讯录').Click()
    pyautogui.moveTo(217,52)
    pyautogui.click()
    ClipboardText(wxid)
    pyautogui.press('enter')
    pyautogui.moveTo(210,143)
    pyautogui.click()
    wx_GetFriend.ButtonControl(Name='聊天信息').Click()
    pyautogui.moveTo(929,64)
    pyautogui.click()
    print('已找到'+wxid+'的微信信息')

def CreateWxNote(NoteContent):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Create = WindowControl(Name='微信')
    wx_Create.ButtonControl(Name='收藏').Click()
    pyautogui.moveTo(209,106)
    pyautogui.click()
    ClipboardText(NoteContent)
    pyautogui.moveTo(1173,181)
    pyautogui.click()
    print('已创建内容为'+NoteContent+'的笔记')

def OpenWxMoments():
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_OpenCircle = WindowControl(Name='微信')
    wx_OpenCircle.ButtonControl(Name='朋友圈').Click()
    print('已打开朋友圈')

def OpenWxVideoId(Videoid):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_OpenVideo = WindowControl(Name='微信')
    wx_OpenVideo.ButtonControl(Name='视频号').Click()
    pyautogui.moveTo(673,132)
    pyautogui.click()
    ClipboardText(Videoid)
    pyautogui.moveTo(527,124)
    pyautogui.press('enter')
    print('已打开'+Videoid+'的视频号')

def OpenWxTakeALook():
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_OpenGlance = WindowControl(Name='微信')
    wx_OpenGlance.ButtonControl(Name='看一看').Click()
    print('已打开看一看')

def OpenWxSearch(Content):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_OpenGlance = WindowControl(Name='微信')
    wx_OpenGlance.ButtonControl(Name='更多功能').Click()
    wx_OpenGlance.ButtonControl(Name='搜一搜').Click()
    pyautogui.moveTo(626,288)
    pyautogui.click()
    ClipboardText(Content)
    pyautogui.press('enter')
    print('已打开搜一搜并搜寻'+Content+'完成')

def OpenWxMiniPrograms(Content):
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_OpenPrograms = WindowControl(Name='微信')
    wx_OpenPrograms.ButtonControl(Name='小程序面板').Click()
    pyautogui.moveTo(1259,122)
    pyautogui.click()
    ClipboardText(Content)
    pyautogui.press('enter')
    print('已打开小程序面板并搜寻'+Content+'成功')

def LockWx():
    FindWindow('微信')
    time.sleep(1)
    # 绑定微信主窗口
    wx_Lock = WindowControl(Name='微信')
    wx_Lock.ButtonControl(Name='设置及其他').Click()
    pyautogui.moveTo(145,534)
    pyautogui.click()
    print('已锁定微信')

def UnlockWx():
    FindWindow('微信')
    time.sleep(1)
    pyautogui.moveTo(438,542)
    pyautogui.click()
    print('已申请解锁微信')
# 制作完成

#                       _oo0oo_                                            
#                      o8888888o                                  
#                      88" . "88  
#                      (| -_- |)
#                      0\  =  /0
#                    ___/`---'\___
#                  .' \\|     |// '.
#                 / \\|||  :  |||// \
#                / _||||| -:- |||||- \
#               |   | \\\  - /// |   |
#               | \_|  ''\---/''  |_/ |
#               \  .-\__  '-'  ___/-. /
#             ___'. .'  /--.--\  `. .'___
#          ."" '<  `.___\_<|>_/___.' >' "".
#         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
#         \  \ `_.   \_ __\ /__ _/   .-` /  /
#     =====`-.____`.___ \_____/___.-`___.-'=====
#                       `=---='
#
#
#     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#           佛祖保佑       永不宕机     永无BUG
  