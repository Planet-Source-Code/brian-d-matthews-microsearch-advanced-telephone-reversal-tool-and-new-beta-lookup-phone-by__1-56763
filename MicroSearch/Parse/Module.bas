Attribute VB_Name = "Module"
Option Explicit
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 6
Public Const MouseMove = &HA1
Public Const Caption = 2
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Function Module_Author()
Module_Author = _
Chr(67) & Chr(111) & Chr(108) & Chr(100) & Chr(32) & _
Chr(66) & Chr(108) & Chr(111) & Chr(111) & _
Chr(100) & Chr(101) & Chr(100) & Chr(32) & Chr(75) & _
Chr(105) & Chr(110) & Chr(103) & vbCrLf & LCase(Chr(72) & _
Chr(84) & Chr(84) & Chr(80) & Chr(58) & Chr(47) & Chr(47) & _
Chr(69) & Chr(45) & Chr(68) & Chr(65) & Chr(82) & Chr(75) & _
Chr(78) & Chr(69) & Chr(83) & Chr(83) & Chr(46) & Chr(67) & _
Chr(74) & Chr(66) & Chr(46) & Chr(78) & Chr(69) & Chr(84))
End Function

Sub RunMenuByString(lngwindow As Long, strmenutext As String)
Dim intLoop As Integer, intSubLoop As Integer, intSub2Loop As Integer, intSub3Loop As Integer, intSub4Loop As Integer
Dim lngmenu(1 To 5) As Long, lngcount(1 To 5) As Long, lngSubMenuID(1 To 4) As Long, strcaption(1 To 4) As String
lngmenu(1) = GetMenu(lngwindow&)
lngcount(1) = GetMenuItemCount(lngmenu(1))
For intLoop% = 0 To lngcount(1) - 1
DoEvents
lngmenu(2) = GetSubMenu(lngmenu(1), intLoop%)
lngcount(2) = GetMenuItemCount(lngmenu(2))
For intSubLoop% = 0 To lngcount(2) - 1
DoEvents
lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop%)
strcaption(1) = String(75, " ")
Call GetMenuString(lngmenu(2), lngSubMenuID(1), strcaption(1), 75, 1)
If InStr(LCase(strcaption(1)), LCase(strmenutext$)) Then
Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(1), 0)
Exit Sub
End If
lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop%)
lngcount(3) = GetMenuItemCount(lngmenu(3))
If lngcount(3) > 0 Then
For intSub2Loop% = 0 To lngcount(3) - 1
DoEvents
lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop%)
strcaption(2) = String(75, " ")
Call GetMenuString(lngmenu(3), lngSubMenuID(2), strcaption(2), 75, 1)
If InStr(LCase(strcaption(2)), LCase(strmenutext$)) Then
Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(2), 0)
Exit Sub
End If
lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop%)
lngcount(4) = GetMenuItemCount(lngmenu(4))
If lngcount(4) > 0 Then
For intSub3Loop% = 0 To lngcount(4) - 1
DoEvents
lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop%)
strcaption(3) = String(75, " ")
Call GetMenuString(lngmenu(4), lngSubMenuID(3), strcaption(3), 75, 1)
If InStr(LCase(strcaption(3)), LCase(strmenutext$)) Then
Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(3), 0)
Exit Sub
End If
lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop%)
lngcount(5) = GetMenuItemCount(lngmenu(5))
If lngcount(5) > 0 Then
For intSub4Loop% = 0 To lngcount(5) - 1
DoEvents
lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop%)
strcaption(4) = String(75, " ")
Call GetMenuString(lngmenu(5), lngSubMenuID(4), strcaption(4), 75, 1)
If InStr(LCase(strcaption(4)), LCase(strmenutext$)) Then
Call SendMessage(lngwindow&, WM_COMMAND, lngSubMenuID(4), 0)
Exit Sub
End If
Next intSub4Loop%
End If
Next intSub3Loop%
End If
Next intSub2Loop%
End If
Next intSubLoop%
Next intLoop%
End Sub

Public Function FilterFilename_FromFullPath(FullPathFileName As String)
Do While InStr(FullPathFileName, "\") <> 0
FullPathFileName = Mid(FullPathFileName, InStr(FullPathFileName, "\") + 1, Len(FullPathFileName) - InStr(FullPathFileName, "\") + 1)
DoEvents
Loop
FilterFilename_FromFullPath = FullPathFileName
End Function

Function Filter_Filename_extension(FileName As String)
Do While InStr(FileName, ".") <> 0
FileName = Mid(FileName, InStr(FileName, ".") + 1, Len(FileName) - InStr(FileName, "\") + 1)
DoEvents
Loop
Filter_Filename_extension = FileName
End Function

Function GetEXE_Path()
GetEXE_Path = App.Path
End Function

Function GetEXE_FileName()
GetEXE_FileName = App.EXEName
End Function

Function GetEXE_PathNfilename()
GetEXE_PathNfilename = App.Path & "\" & App.EXEName
End Function

Function GetEXE_Title()
GetEXE_Title = App.Title
End Function

Sub Remove_Duplicate(ListBx As listbox)
Dim x
Do
ListBx.text = ListBx.list(x)
If Not ListBx.ListIndex = x Then ListBx.RemoveItem x
If ListBx.ListIndex = x Then x = x + 1
Loop Until x > ListBx.ListCount - 1
ListBx.ListIndex = 0
ListBx.text = ""
End Sub

Sub Remove_Duplicate2(ListBx As listbox)
Dim z As Integer, x As Integer
For z = 0 To ListBx.ListCount - 1
For x = 0 To ListBx.ListCount - 1
If Not x = z Then
If ListBx.list(z) = ListBx.list(x) Then
ListBx.RemoveItem x
x = x - 1
End If: End If
Next: Next
End Sub

Sub Pause(interval)
Dim x
x = Timer
Do While Timer - x < Val(interval)
DoEvents
Loop
End Sub

Public Sub AddAscii(Control As Control)
Dim x As Long
Control.Font = "Arial"
For x = 123 To 255
Control.AddItem Chr(x)
DoEvents
Next
End Sub

Public Sub AddFonts(Control As Control)
Dim x As Integer
For x = 1 To Screen.FontCount
Control.AddItem Screen.Fonts(x)
Next x
End Sub

Sub Scrll_ListBx_Bottom(ListBx As listbox)
ListBx.ListIndex = ListBx.ListCount - 1
ListBx.text = ""
End Sub

Sub Run_Notepad()
Shell "notepad.exe", vbNormalFocus
End Sub

Sub Run_Calculator()
Shell "calc.exe", vbNormalFocus
End Sub

Sub Run_Paint()
Shell "mspaint.exe", vbNormalFocus
End Sub

Sub Run_SoundRecorder()
Shell "sndrec32.exe", vbNormalFocus
End Sub

Sub Run_VolumeControl()
Shell "sndvol32.exe", vbNormalFocus
End Sub

Sub Run_CDPlayer()
Shell "cdplayer.exe", vbNormalFocus
End Sub

Sub Run_Disk_Cleanup()
Shell "cleanmgr.exe", vbNormalFocus
End Sub

Sub GoToWebsite(Url As String)
ShellExecute &O0, "Open", Url, vbNullString, vbNullString, SW_NORMAL
End Sub

Sub ViewSitesSource(Url As String)
If Left(LCase(Url), 7) <> "http://" Then Url = "http://" & Url
Shell "Explorer view-source:" & Url
End Sub

Sub Taskbar(Visible As Boolean)
'To Hide: Taskbar False
'To Show: Taskbar True
Dim x As Long
x = FindWindow("Shell_traywnd", "")
If Visible = True Then Call SetWindowPos(x, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
If Visible = False Then Call SetWindowPos(x, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Public Sub TextBx_SelectAll(TextBox As TextBox)
TextBox.SelStart = 0
TextBox.SelLength = Len(TextBox)
TextBox.SetFocus
End Sub

Sub Textbox_Clear(text As TextBox)
text.text = ""
End Sub

Sub Listbox_Clear(list As listbox)
list.Clear
End Sub

Sub Clear_Clipboar()
Clipboard.Clear
End Sub

Sub Copy_Text_to_Clipboard(txt As String)
Clipboard.Clear
Clipboard.SetText txt
End Sub

Function Get_Caption()
Dim imclass As Long, text As String, TLn As Long
imclass = FindWindow("imclass", vbNullString)
TLn = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
text = String(TLn + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TLn + 1, text)
Get_Caption = Left(text, TLn)
End Function

Sub Open_File(Frm As Form, FullPathFileName As String)
On Error GoTo hell
ShellExecute Frm.hwnd, "open", FullPathFileName, vbNullString, vbNullString, 1
Exit Sub
hell: MsgBox Err.Description, vbExclamation
End Sub

Sub Copy_List_to_List(Fromlist As listbox, ToList As listbox)
Dim x As Integer
If Fromlist.ListCount = 0 Then Exit Sub
For x = 0 To Fromlist.ListCount - 1
ToList.AddItem Fromlist.list(x)
DoEvents
Next
End Sub

Sub TextBox_Save(Txtbx As TextBox, FileName As String)
On Error GoTo hell
Open FileName For Output As #1
Print #1, Txtbx.text
Close #1
Exit Sub
hell: MsgBox Err.Description, vbExclamation, "Error"
End Sub

Sub TextBox_Open(Txtbx As TextBox, FileName As String)
On Error GoTo hell
Dim x As String
Open FileName For Input As #1
While Not EOF(1)
Txtbx.text = Input(LOF(1), #1)
Wend
Close #1
Exit Sub
hell: MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function Line_Count_TxtFile(FileName As String)
'can use this to count how many Passwords in the PW List
On Error GoTo hell
Dim o As String, x
x = 0
Open FileName For Input As #1
While Not EOF(1)
Input #1, o
x = x + 1
DoEvents
Wend
Close #1
Line_Count_TxtFile = x
Exit Function
hell: MsgBox Err.Description, vbExclamation, "Error"
End Function

Sub Set_Wind_Caption(Wind As Long, Caption As String)
SendMessageByString Wind, WM_SETTEXT, 0, Caption
End Sub

Function Get_Wind_Caption(Wind As Long)
Dim text As String, TLn As Long
TLn = SendMessageLong(Wind, WM_GETTEXTLENGTH, 0&, 0&)
text = String(TLn + 1, " ")
Call SendMessageByString(Wind, WM_GETTEXT, TLn + 1, text)
Get_Wind_Caption = Left(text, TLn)
End Function
