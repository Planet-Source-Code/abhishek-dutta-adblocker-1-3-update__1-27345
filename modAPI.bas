Attribute VB_Name = "modAPI"
'
Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type
'
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
'
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
'
Public Const LB_DELETESTRING = &H182
'Public Const HWND_TOP = 0
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_SHOWWINDOW = &H40
'
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
'
Public Const WM_CLOSE = &H10
'
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
'
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
Public Declare Function DeleteUrlCacheEntry Lib "wininet" (ByVal lpFileName As String) As Boolean
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'


Public Sub LoadFileLst(box As ListBox, file As String, Optional prefix As String)

Dim found As Boolean
found = False

On Error GoTo err

'load file into listbox
box.clear

Dim row As String

Close #99

Open file For Input As #99

While Not EOF(99)
    Line Input #99, row
    If Len(Trim(row)) > 0 Then AddStringToList Trim((prefix + row)), box
Wend

Close #99

Exit Sub

err:
Call LogError(err)
End Sub

Public Sub AppendFileLst(box As ListBox, file As String)

On Error GoTo err

'add contents of a text box to a file
Close #99

Open file For Append As #99
    For i = 0 To box.ListCount - 1
        Print #99, Trim(box.List(i))
    Next
    Print #99, Chr(13) + Chr(10)
Close #99
box.clear

Exit Sub

err:
Call LogError(err)
End Sub

Public Sub SaveFileLst(box As ListBox, file As String, Optional clear As Boolean)
'save listbox contents as file
Close #99

Open file For Output As #99
    For i = 0 To box.ListCount - 1
        Print #99, Trim(box.List(i))
    Next
    Print #99, Chr(13) + Chr(10)
Close #99

If clear <> False Then box.clear
End Sub

Public Sub CreateBlankFile(file As String)
'create a blank file if it dows not exist

If Dir(file) = "" Then
    frmMain.lstFile.clear
    Call SaveFileLst(frmMain.lstFile, file)
End If

End Sub

Public Sub LogError(err As ErrObject)
frmMain.lstFile.clear
frmMain.lstFile.AddItem Str(err.Number) + ":" + err.Description + ":" + err.Source
Call CreateBlankFile(App.Path + "\err.log")
Call AppendFileLst(frmMain.lstFile, App.Path + "\err.log")
frmMain.lstFile.clear
End Sub

Public Function GetUName() As String
Dim r As Long
Dim strUserName As String * 255
Dim strRet As String

r = GetUserName(strUserName, 255)
strUserName = Left(strUserName, InStr(strUserName, Chr(0)) - 1)

strRet = Trim(strUserName)
GetUName = strRet
End Function
Public Function AddStringToList(strinput As String, box As ListBox, Optional lstsetfocus As Boolean)
Dim i As Integer
For i = 0 To box.ListCount
    If strinput = box.List(i) Then
        box.ListIndex = i
        AddStringToList = False
        Exit Function
    End If
Next

box.AddItem strinput

End Function
