VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About AdBlocker v 1.3"
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   2205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstpops 
      Height          =   1425
      Left            =   3825
      TabIndex        =   7
      Top             =   2175
      Width           =   1050
   End
   Begin VB.ListBox lstFile 
      Height          =   1425
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   540
      TabIndex        =   2
      Top             =   1290
      Width           =   1125
   End
   Begin VB.FileListBox fileCk 
      Height          =   1455
      Left            =   1305
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   2160
      Width           =   1140
   End
   Begin VB.ListBox lstCk 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   795
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   135
      Top             =   3690
   End
   Begin VB.Image imgicon 
      Height          =   480
      Left            =   210
      Picture         =   "frmMain.frx":0442
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "AdBlocker v 1.3"
      Height          =   180
      Left            =   765
      TabIndex        =   5
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Abhishek Dutta"
      Height          =   210
      Left            =   127
      TabIndex        =   4
      Top             =   675
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "<abhishekdutta@mail.com>"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   105
      MouseIcon       =   "frmMain.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   930
      Width           =   1995
   End
   Begin VB.Menu mnuconfigure 
      Caption         =   "&Configure"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnublockads 
         Caption         =   "Block &Ads"
      End
      Begin VB.Menu mnukillcookies 
         Caption         =   "Kill &Cookies"
      End
      Begin VB.Menu mnublockpopups 
         Caption         =   "Block &Popups"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NID As NOTIFYICONDATA
Dim StrWinDir As String
Dim Status As Boolean
Dim Stats(3) As Integer
Dim Hnd As Long
Dim Hcld As Long
Dim Class As String * 16
Dim cons As Integer
'

Private Sub cmdOk_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Visible = False

'terminate if already running
If App.PrevInstance Then
    End
End If

CheckFiles

On Error GoTo err

'declare / initialize variables
Dim r As Integer
Dim struname As String

StrWinDir = Space(255)
cons = 0
For i = 0 To 3
    Stats(i) = 0
Next

' get windows directory name
r = GetWindowsDirectory(StrWinDir, Len(StrWinDir))

If r = 0 Then
    MsgBox "Error retrieveing Windows directory name" + vbNewLine + "C:\WINDOWS will be used as the default windows directory", vbCritical + vbOKOnly, "Error"
    StrWinDir = "C:\WINDOWS" 'on error use default
Else
    StrWinDir = Left(StrWinDir, r)
End If

' get the user name
struname = GetUName()

'get location of windows hosts file and cookies diretory
'based on the os version
If GetVersion() = 0 Then
    'if nt then set appropriate paths for cookie dir and hosts file
    fileCk.Path = (StrWinDir + "\profiles\" + struname + "\" + "cookies")
    StrWinDir = (StrWinDir + "\system32\drivers\etc\hosts")
Else
    'if win95/98 etc then set apt paths
    If Len(struname) > 0 Then
         'if user is logged in the cookies dir is \profiles\username\cookies
        If Dir((StrWinDir + "\profiles\" + struname + "\cookies"), vbDirectory Or vbSystem) <> "" Then fileCk.Path = (StrWinDir + "\profiles\" + struname + "\cookies")
     Else
         'else normal cookies dir \cookies
        fileCk.Path = StrWinDir + "\cookies"
     End If
     
     StrWinDir = (StrWinDir + "\hosts")
    
End If

' init systray icon etc. params & add icon to systray
With NID
 .cbSize = Len(NID)
 .hwnd = frmMain.hwnd
 .uId = vbNull
 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 .uCallBackMessage = WM_MOUSEMOVE
 .hIcon = frmMain.Icon
 .szTip = "AdBlocker 1.3" & vbNullChar
End With

'start our own ad server!
Call StartServer
    
'add icon to systray
Shell_NotifyIcon NIM_ADD, NID

'set attributes of files to vbnormal
Call SetAttributes

'back up original hosts file, only if it was not restored properly the last time
If GetSetting("AdBlocker", "Data", "Restore", "True") = "True" Then
    CreateBlankFile (StrWinDir) 'create hosts if it does not exists
    Call LoadFileLst(frmMain.lstFile, StrWinDir)
    Call SaveFileLst(frmMain.lstFile, App.Path + "\data\hosts.bak")
    'hosts file changed, store it in registry
    Call SaveSetting("AdBlocker", "Data", "Restore", "False")
End If
 
 
'create log files if they dont' exist
Call CreateBlankFile(App.Path + "\ads.log")
Call CreateBlankFile(App.Path + "\cookies.log")

'check and delete large log files
If FileLen(App.Path + "\ads.log") > 16000 Then Call Kill(App.Path + "\ads.log")
If FileLen(App.Path + "\cookies.log") > 16000 Then Call Kill(App.Path + "\cookies.log")

'get previous settings if any
Call InitSettings

Exit Sub

err:

MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)
Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Stats(3) = X / Screen.TwipsPerPixelX

    Select Case Stats(3)
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
            Call mnusettings_Click
       Case WM_RBUTTONDOWN
           'set this form as current window, otherwise popup
           'menu will stay on the screen even after you click
           'outside it
           Call SetForegroundWindow(Me.hwnd)
           PopupMenu frmMain.mnuconfigure
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
'show stats
'MsgBox "Ads Blocked : " + Str(stats(0)) + vbNewLine + "Cookies Killed : " + Str(stats(1)), vbInformation, "Statistics"
If frmSettings.Visible = True Then Unload frmSettings

'stop adblocking, baically restore the hosts file
blockads (False)

Call StopServer

'save current settings
Call SaveSettings

'remove icon from systray
Shell_NotifyIcon NIM_DELETE, NID

End

End Sub

Private Sub Label3_Click()
'don't forget to send me a mail!
ShellExecute 0, "open", "mailto:abhishekdutta@mail.com", vbNullString, vbNullString, vbshownormal
End Sub

Private Sub mnuabout_Click()
Me.Visible = True
End Sub

Private Sub mnublockads_Click()
mnublockads.Checked = Not mnublockads.Checked

'turn on/off adblocking
Call blockads(mnublockads.Checked)
End Sub

Private Sub mnublockpopups_Click()
mnublockpopups.Checked = Not mnublockpopups.Checked

'turn on/off popup blocking
Call killpopups(mnublockpopups.Checked)
End Sub

Private Sub mnuexit_Click()
Call Form_Unload(False)
End Sub


Public Sub mnukillcookies_Click()
mnukillcookies.Checked = Not mnukillcookies.Checked

'turn on/off cookie killing
Call KillCookies(mnukillcookies.Checked)
End Sub

Private Sub mnusettings_Click()
'
Load frmSettings
frmSettings.Show

End Sub

Private Sub Timer_Timer()

Class = ""
Hnd = 0
Hcld = 0
Status = False

On Error GoTo err:

'check active window
If mnublockpopups.Checked = True Then
    Hnd = GetForegroundWindow()
    Call GetClassName(Hnd, Class, 15)
    'if the window is an IE window then
    If Left(Class, 7) = "IEFrame" Or Left(Class, 13) = "CabinetWClass" Then

        'if parent ie window is visible then
        If IsWindowVisible(Hnd) <> False Then
        
            'find first child window
            Hcld = FindWindowEx(Hnd, GW_CHILD And GW_HWNDFIRST, vbNullString, vbNullString)
            Call GetClassName(Hcld, Class, 8)
            
            'if WorkerA (ie animation on top right) class is found and it is invisible
            'then the window is a popup
            If Left(Class, 7) = "WorkerA" And IsWindowVisible(Hcld) = False Then
                If lstpops.ListCount = 0 Then
                    Call PostMessage(Hnd, WM_CLOSE, 0, 0)
                    Call Sleep(100)
                    Stats(2) = Stats(2) + 1
                Else
                    
                    'check whether should it be closed or not
                                         
                    Call Sleep(500)
                    Dim cap As String * 100
                    Call GetWindowText(Hnd, cap, 100)
                        
                    For j = 0 To lstpops.ListCount - 1
                        If InStr(1, cap, lstpops.List(j), vbTextCompare) > 0 Then
                            Status = True 'dont close this window because user doesn't want so
                            Exit For
                        End If
                    Next
                        
                                  
                    If Status = False Then
                    'close the popup
                        Call PostMessage(Hnd, WM_CLOSE, 0, 0)
                        Stats(2) = Stats(2) + 1
                    End If
                    
                End If
            End If
        End If
    End If
End If
    
If mnukillcookies.Checked = True Then
    fileCk.Refresh

    'delete first cookie if selective cookie deletion is disabled
    If lstCk.ListCount = 0 Then
        Call KillCookie(fileCk.List(0))
    Else
    'find out files names of cookies to be deleted and delete them
        For i = 0 To fileCk.ListCount - 1
            Status = False
    
                For j = 0 To lstCk.ListCount - 1
        
                    If InStr(1, fileCk.List(i), lstCk.List(j), vbTextCompare) > 1 Then
                        Status = True
                        Exit For
                    End If
                Next
    
            If Status = False Then
                Call KillCookie(fileCk.List(i))
            End If
 
        Next
    End If
End If

NID.szTip = "Ads:" + Str(Stats(0)) + " Cookies:" + Str(Stats(1)) + " Popups:" + Str(Stats(2)) + vbNullChar
Shell_NotifyIcon NIM_MODIFY, NID

Exit Sub

err:
Call LogError(err)
Exit Sub

End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error GoTo err

'if connection request is not from local system reject it
If wsock(0).RemoteHostIP = "127.0.0.1" Then
    cons = cons + 1
    Load wsock(cons)
    wsock(cons).LocalPort = 0
    wsock(cons).Accept requestID
End If

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo err

Dim strdata As String
Dim Page As String

strdata = ""

' if the browser request for any ads, send it an appropriate image
wsock(Index).GetData strdata$

If Mid$(strdata$, 1, 3) = "GET" Then
    
 
    'store inforamtion in a log file
    Call CreateBlankFile(App.Path + "\ads.log")
    strdata = Left(strdata, (InStr(1, strdata, "HTTP") - 2))
    strdata = Mid(strdata, 4)
    lstFile.AddItem strdata + " send on " + Str(Now())
    Call AppendFileLst(frmMain.lstFile, App.Path + "\ads.log")
    
    wsock(cons).SendData SendPageStr(Trim(StrConv(strdata, vbLowerCase)))
Else
    'close the socket if request is not valid
    wsock(cons).Close
End If

Exit Sub

err:

'MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub wsock_SendComplete(Index As Integer)
wsock(cons).Close
Unload wsock(cons)
End Sub

Private Sub StartServer()
On Error GoTo err

Connections = 0
Me.wsock(0).Close
Me.wsock(0).Bind "80", "127.0.0.1"
Me.wsock(0).Listen

Exit Sub

err:

Call LogError(err)

If err.Number = 10048 Then
    MsgBox "A webserver is already running on port 80, for best results please stop the server and restart this application", vbOKOnly, "Error"
    Form_Unload (False)
    End
Else
    MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
End If

End Sub

Private Sub StopServer()
wsock(0).Close
End Sub

Private Sub SaveSettings()
On Error GoTo err

Call SaveSetting("AdBlocker", "Data", "AdBlock", mnublockads.Checked)
Call SaveSetting("AdBlocker", "Data", "Cookies", mnukillcookies.Checked)
Call SaveSetting("AdBlocker", "Data", "Popup", mnublockpopups.Checked)
Call SaveSetting("AdBlocker", "Data", "Restore", "True")

Call SetAttr(App.Path + "\data\hosts.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\cookies.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\dns.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\popups.svr", vbReadOnly + vbSystem)
Call SetAttr(StrWinDir, vbReadOnly)

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub InitSettings()
On Error GoTo err

If GetSetting("AdBlocker", "Data", "AdBlock", False) = True Then Call mnublockads_Click
If GetSetting("AdBlocker", "Data", "Cookies", False) = True Then Call mnukillcookies_Click
If GetSetting("AdBlocker", "Data", "Popup", False) = True Then Call mnublockpopups_Click

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub KillCookie(CookieName)
If CookieName = "" Then Exit Sub
Dim i As Integer, j As Integer
Dim temp As String

On Error Resume Next

'get the name of the server which created the cookie
Call LoadFileLst(frmMain.lstFile, fileCk.Path + "\" + CookieName)

temp = lstFile.List(0)

i = 0
j = 0

'read server name from the cookie
While (i < 2)
        j = InStr(1, temp, Chr(10), vbBinaryCompare)
        temp = Right(temp, Len(temp) - j)
        i = i + 1
Wend

j = InStr(1, temp, Chr(10), vbBinaryCompare)
temp = Left(temp, j - 1)

'get the cache entry
If Len(GetUName()) = 0 Then
    temp = "Cookie:anyuser@" + temp
Else
    temp = "Cookie:" + StrConv(GetUName(), vbLowerCase) + "@" + temp
End If

'clear the listbox
lstFile.clear

Call DeleteFile(fileCk.Path + "\" + CookieName)

'remove the cookie cache entry
Call DeleteUrlCacheEntry(temp)

'save info in log file
Call CreateBlankFile(App.Path + "\cookies.log")
lstFile.AddItem (fileCk.Path + "\" + CookieName + " deleted on " + Str(Now()))
Call AppendFileLst(frmMain.lstFile, App.Path + "\cookies.log")
           
'increment the count
Stats(1) = Stats(1) + 1

End Sub

Private Sub SetAttributes()

On Error GoTo err

Call SetAttr(App.Path + "\data\hosts.svr", vbNormal)
Call SetAttr(App.Path + "\data\cookies.svr", vbNormal)
Call SetAttr(App.Path + "\data\dns.svr", vbNormal)
Call SetAttr(App.Path + "\data\popups.svr", vbNormal)
If Dir(StrWinDir, vbHidden Or vbSystem Or vbReadOnly) <> "" Then Call SetAttr(StrWinDir, vbNormal)
Exit Sub

err:

MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Function SendPageStr(pType As String) As String

Dim tx As String
Dim tx1 As String
Dim lg As Integer
Dim bgcolor As String
On Error GoTo err

tx = " "
tx1 = ""

'get settings - color of ad basically
bgcolor = Trim(GetSetting("AdBlocker", "Settings", "Color", "0"))


If pType Like "*.htm*" Or pType Like "*.asp*" Or pType Like "*/" Or pType Like "*ord=*" Then
    
     Select Case bgcolor
     Case 0
        bgcolor = ""
     Case 1
        bgcolor = "#FFFFFF"
     Case 2
        bgcolor = "#0000FF"
     Case 3
        bgcolor = "#00FF00"
     Case 4
        bgcolor = "#FF0000"
     Case 5
        bgcolor = "#000000"
     End Select
     
     pType = "text/html"
     tx1 = "<html><body bgcolor=" + bgcolor + ">Ad blocked by AdBlocker</body></html>"
     lg = 84
ElseIf pType Like "*.js*" Then
     
     pType = "text/javascript"
     tx1 = "document.write(" + Chr(34) + "<table width=100% background=http://127.0.0.1/bg.gif><tr><td>Ad blocked by AdBlocker</td></tr></table>" + Chr(34) + ")"
     lg = 117
Else
     bgcolor = App.Path + "\images\" + Trim(bgcolor) + ".gif"
     Nr = FreeFile
     lg = FileLen(bgcolor)
     pType = "image/gif"
     
     Open (bgcolor) For Binary As Nr

         For m = 1 To lg
             Get #Nr, , tx
             tx1 = tx1 + tx
         Next
     Close Nr
End If

Stats(0) = Stats(0) + 1

SendPageStr = "HTTP/1.1 200 OK" + vbNewLine + _
"Expires: 0" + vbNewLine + _
"Last-Modified: 0" + vbNewLine + _
"Accept -Range: bytes" + vbNewLine + _
"Content-Type: " + pType + vbNewLine + _
"Content-Length: " + Str(lg) + vbNewLine + vbNewLine + _
tx1

Exit Function
err:

End Function
Public Sub killpopups(state As Boolean)

lstpops.clear

If state = True Then
    If CInt(GetSetting("AdBlocker", "Settings", "AllPops", 1)) = 0 Then Call LoadFileLst(lstpops, App.Path + "\data\popups.svr")
End If



End Sub


Public Sub KillCookies(state As Boolean)
On Error GoTo err:

lstCk.clear
    
'start cookie killing, but before that get names of servers from where
'cookies should not be deleted
If state = True Then
    If CInt(GetSetting("AdBlocker", "Settings", "All", "0")) = 0 Then
        Call LoadFileLst(lstCk, App.Path + "\data\cookies.svr")
    End If
End If

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Public Sub blockads(state As Boolean)

On Error GoTo err:

If state = True Then
    
    ' if adblocking is enabled, load contents from hosts.svr
    'and save it to system's hosts file
    Call LoadFileLst(frmMain.lstFile, App.Path + "\data\hosts.svr", "127.0.0.1 ")
    Call SaveFileLst(frmMain.lstFile, StrWinDir)
    
    ' if dns caching is enabled, load dns,svr and append to hosts file
    If GetSetting("AdBlocker", "Settings", "UseDNS", "0") = "1" Then
        Call LoadFileLst(frmMain.lstFile, App.Path + "\data\dns.svr")
        Call AppendFileLst(frmMain.lstFile, StrWinDir)
    End If
    
Else
    ' restore the contents of the original hosts file
    Call LoadFileLst(frmMain.lstFile, App.Path + "\data\hosts.bak")
    Call SaveFileLst(frmMain.lstFile, StrWinDir)
End If


Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Function CheckFiles()

If Dir(App.Path + "\data\hosts.svr", vbReadOnly Or vbSystem) = "" Or Dir(App.Path + "\data\dns.svr", vbReadOnly Or vbSystem) = "" Or Dir(App.Path + "\data\cookies.svr", vbReadOnly Or vbSystem) = "" Or Dir(App.Path + "\data\popups.svr", vbReadOnly Or vbSystem) = "" Then
    MsgBox "The required settings file(s) are missing. Please reinstall AdBlocker. You may also download the files from http://germany.ms/abhishek", vbCritical, "Error"
    End
End If

End Function
