VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSettings.frx":0442
   ScaleHeight     =   4050
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ws 
      Left            =   750
      Top             =   4395
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4980
      TabIndex        =   37
      Top             =   3555
      Width           =   1005
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3780
      TabIndex        =   36
      Top             =   3555
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok "
      Height          =   345
      Left            =   2520
      TabIndex        =   35
      Top             =   3555
      Width           =   1005
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   3315
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5847
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   0
      TabCaption(0)   =   "Ad&Blocking"
      TabPicture(0)   =   "frmSettings.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbladserver"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbocols"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lstadservers"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtadserver"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAddServer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdRmvServer"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Cookies"
      TabPicture(1)   =   "frmSettings.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "lblcookies"
      Tab(1).Control(3)=   "lstcookies"
      Tab(1).Control(4)=   "txtcookies"
      Tab(1).Control(5)=   "cmdAddCookies"
      Tab(1).Control(6)=   "cmdRmvCookies"
      Tab(1).Control(7)=   "chkallcookies"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&DNS Caching"
      TabPicture(2)   =   "frmSettings.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstdns"
      Tab(2).Control(1)=   "txtdns"
      Tab(2).Control(2)=   "cmdAddDns"
      Tab(2).Control(3)=   "cmdRmvDns"
      Tab(2).Control(4)=   "chkusedns"
      Tab(2).Control(5)=   "lbldns"
      Tab(2).Control(6)=   "Label8"
      Tab(2).Control(7)=   "Label7"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "&Popups"
      TabPicture(3)   =   "frmSettings.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(2)=   "lblpopups"
      Tab(3).Control(3)=   "lstpopups"
      Tab(3).Control(4)=   "txtpopups"
      Tab(3).Control(5)=   "cmdrmvpops"
      Tab(3).Control(6)=   "cmdaddpops"
      Tab(3).Control(7)=   "chkblockallpops"
      Tab(3).ControlCount=   8
      Begin VB.CheckBox chkblockallpops 
         Caption         =   "Block all popups"
         Height          =   210
         Left            =   -74790
         TabIndex        =   30
         Top             =   555
         Width           =   1605
      End
      Begin VB.CommandButton cmdaddpops 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -71205
         TabIndex        =   3
         Top             =   2835
         Width           =   870
      End
      Begin VB.CommandButton cmdrmvpops 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   -70200
         TabIndex        =   4
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtpopups 
         Height          =   285
         Left            =   -73890
         TabIndex        =   2
         Top             =   2850
         Width           =   2550
      End
      Begin VB.ListBox lstpopups 
         Height          =   1620
         Left            =   -74790
         TabIndex        =   1
         Top             =   1110
         Width           =   5460
      End
      Begin VB.ListBox lstdns 
         Height          =   1620
         Left            =   -74790
         TabIndex        =   6
         Top             =   1110
         Width           =   5460
      End
      Begin VB.TextBox txtdns 
         Height          =   285
         Left            =   -73890
         TabIndex        =   7
         Top             =   2843
         Width           =   2550
      End
      Begin VB.CommandButton cmdAddDns 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -71205
         TabIndex        =   8
         Top             =   2835
         Width           =   870
      End
      Begin VB.CommandButton cmdRmvDns 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   -70200
         TabIndex        =   9
         Top             =   2835
         Width           =   870
      End
      Begin VB.CheckBox chkusedns 
         Caption         =   "Use DNS Caching"
         Height          =   210
         Left            =   -74790
         TabIndex        =   5
         Top             =   555
         Width           =   1605
      End
      Begin VB.CheckBox chkallcookies 
         Caption         =   "Delete all cookies"
         Height          =   210
         Left            =   -74805
         TabIndex        =   10
         ToolTipText     =   "Delete all cookies from the cookies folder"
         Top             =   555
         Width           =   1605
      End
      Begin VB.CommandButton cmdRmvCookies 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   -70200
         TabIndex        =   14
         Top             =   2835
         Width           =   870
      End
      Begin VB.CommandButton cmdAddCookies 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -71205
         TabIndex        =   13
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtcookies 
         Height          =   285
         Left            =   -73890
         TabIndex        =   12
         Top             =   2850
         Width           =   2550
      End
      Begin VB.ListBox lstcookies 
         Height          =   1620
         Left            =   -74790
         TabIndex        =   11
         ToolTipText     =   "Tcookies frm these servers are not to be deleted"
         Top             =   1110
         Width           =   5460
      End
      Begin VB.CommandButton cmdRmvServer 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   4800
         TabIndex        =   19
         Top             =   2835
         Width           =   870
      End
      Begin VB.CommandButton cmdAddServer 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3795
         TabIndex        =   18
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtadserver 
         Height          =   285
         Left            =   1110
         TabIndex        =   17
         Top             =   2850
         Width           =   2550
      End
      Begin VB.ListBox lstadservers 
         Height          =   1620
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "List of blocked ad servers"
         Top             =   1110
         Width           =   5460
      End
      Begin VB.ComboBox cbocols 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Select the color with which you want to block the ad"
         Top             =   495
         Width           =   1455
      End
      Begin VB.Label lblpopups 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   -70470
         TabIndex        =   34
         Top             =   885
         Width           =   1140
      End
      Begin VB.Label lbldns 
         Alignment       =   1  'Right Justify
         Height          =   165
         Left            =   -70620
         TabIndex        =   33
         Top             =   885
         Width           =   1290
      End
      Begin VB.Label lblcookies 
         Alignment       =   1  'Right Justify
         Height          =   165
         Left            =   -70350
         TabIndex        =   32
         Top             =   885
         Width           =   1020
      End
      Begin VB.Label lbladserver 
         Alignment       =   1  'Right Justify
         Height          =   165
         Left            =   4620
         TabIndex        =   31
         Top             =   885
         Width           =   1050
      End
      Begin VB.Label Label10 
         Caption         =   "Add to list"
         Height          =   195
         Left            =   -74805
         TabIndex        =   29
         Top             =   2910
         Width           =   840
      End
      Begin VB.Label Label9 
         Caption         =   "Don't block popups with these captions"
         Height          =   195
         Left            =   -74790
         TabIndex        =   28
         Top             =   885
         Width           =   2910
      End
      Begin VB.Label Label8 
         Caption         =   "Cached Addresses"
         Height          =   165
         Left            =   -74790
         TabIndex        =   27
         Top             =   885
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Add to List"
         Height          =   165
         Left            =   -74805
         TabIndex        =   26
         Top             =   2903
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "Add to List"
         Height          =   165
         Left            =   -74805
         TabIndex        =   25
         Top             =   2910
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Allow cookies from these sites"
         Height          =   165
         Left            =   -74790
         TabIndex        =   24
         Top             =   885
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "Add to List"
         Height          =   165
         Left            =   195
         TabIndex        =   23
         Top             =   2910
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "List of blocked ad servers"
         Height          =   165
         Left            =   210
         TabIndex        =   22
         Top             =   885
         Width           =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "image"
         Height          =   210
         Left            =   3120
         TabIndex        =   21
         Top             =   547
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Replace ads with "
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   555
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbocols_Click()
Call EnableApply
End Sub

Private Sub chkallcookies_Click()
Call EnableApply
lstcookies.Enabled = Not chkallcookies.Value
End Sub

Private Sub chkblockallpops_Click()
Call EnableApply
End Sub

Private Sub chkusedns_Click()
Call EnableApply
Call CheckDNS
End Sub

Private Sub cmdAddCookies_Click()
If AddStringToList(Trim(txtcookies.Text), frmSettings.lstcookies) = True Then EnableApply
Call SetLstFocus(frmSettings.lstcookies)
txtcookies.Text = ""
lblcookies.Caption = Str(lstcookies.ListCount)
End Sub

Private Sub cmdAddDns_Click()
ws.Close
ws.Connect Trim(txtdns.Text), "80"
End Sub

Private Sub cmdaddpops_Click()
If AddStringToList(Trim(txtpopups.Text), frmSettings.lstpopups) = True Then Call EnableApply
Call SetLstFocus(frmSettings.lstpopups)
txtpopups.Text = ""
lblpopups.Caption = Str(lstpopups.ListCount)

End Sub

Private Sub cmdAddServer_Click()
If AddStringToList(Trim(txtadserver.Text), frmSettings.lstadservers) = True Then Call EnableApply
Call SetLstFocus(frmSettings.lstadservers)
txtadserver.Text = ""
lbladserver.Caption = Str(lstadservers.ListCount)
End Sub

Private Sub cmdapply_Click()
Call SaveFileLst(lstadservers, App.Path + "\data\hosts.svr", False)
Call SaveFileLst(lstcookies, App.Path + "\data\cookies.svr", False)
Call SaveFileLst(lstdns, App.Path + "\data\dns.svr", False)
Call SaveFileLst(lstpopups, App.Path + "\data\popups.svr", False)
Call SaveSetting("AdBlocker", "Settings", "Color", Str(cbocols.ListIndex))
Call SaveSetting("AdBlocker", "Settings", "All", chkallcookies.Value)
Call SaveSetting("AdBlocker", "Settings", "UseDNS", chkusedns.Value)
Call SaveSetting("AdBlocker", "Settings", "AllPops", chkblockallpops.Value)

With frmMain
    .blockads (.mnublockads.Checked)
    .KillCookies (.mnukillcookies.Checked)
    .killpopups (.mnublockpopups.Checked)
End With

cmdapply.Enabled = False

End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Call cmdapply_Click
Call cmdcancel_Click
End Sub

Private Sub cmdRmvCookies_Click()
If lstcookies.ListIndex = -1 Then Exit Sub
lstcookies.RemoveItem lstcookies.ListIndex

Call EnableApply

txtcookies.Text = ""
lblcookies.Caption = Str(lstcookies.ListCount)

End Sub

Private Sub cmdRmvDns_Click()
If lstdns.ListIndex = -1 Then Exit Sub
lstdns.RemoveItem lstdns.ListIndex

Call EnableApply

txtdns.Text = ""
lbldns.Caption = Str(lstdns.ListCount)
End Sub

Private Sub cmdrmvpops_Click()
If lstpopups.ListIndex = -1 Then Exit Sub
lstpopups.RemoveItem lstpopups.ListIndex

Call EnableApply

txtpopups.Text = ""
lblpopups.Caption = Str(lstpopups.ListCount)
End Sub

Private Sub cmdRmvServer_Click()
If lstadservers.ListIndex = -1 Then Exit Sub
lstadservers.RemoveItem lstadservers.ListIndex

Call EnableApply

txtadserver.Text = ""
lbladserver.Caption = Str(lstadservers.ListCount)

End Sub


Private Sub Form_Load()
With cbocols
    .AddItem "Transparent"
    .AddItem "White"
    .AddItem "Blue"
    .AddItem "Green"
    .AddItem "Red"
    .AddItem "Black"
    .ListIndex = 0
End With

Call initFTab
Call InitSettings
Call CheckDNS
Call CheckAllCookies
cmdapply.Enabled = False

Call LoadFileLst(lstcookies, App.Path + "\data\cookies.svr")
lblcookies.Caption = Str(lstcookies.ListCount)
Call LoadFileLst(lstdns, App.Path + "\data\dns.svr")
lbldns.Caption = Str(lstdns.ListCount)
Call LoadFileLst(lstpopups, App.Path + "\data\popups.svr")
lblpopups.Caption = Str(lstpopups.ListCount)

End Sub

Private Sub lstadservers_Click()
On Error Resume Next
txtadserver.Text = lstadservers.List(lstadservers.ListIndex)
End Sub

Private Sub lstcookies_Click()
txtcookies.Text = lstcookies.List(lstcookies.ListIndex)
End Sub

Private Sub lstdns_Click()
txtdns.Text = lstdns.List(lstdns.ListIndex)
End Sub

Private Sub lstpopups_Click()
txtpopups.Text = lstpopups.List(lstpopups.ListIndex)
End Sub

Private Sub txtadserver_Change()
cmdAddServer.Enabled = Len(Trim(txtadserver.Text))
End Sub

Private Sub txtcookies_Change()
cmdAddCookies.Enabled = Len(Trim(txtcookies.Text))
End Sub

Private Sub txtdns_Change()
cmdAddDns.Enabled = Len(Trim(txtdns.Text))
End Sub

Private Sub InitSettings()
cbocols.ListIndex = CInt(GetSetting("AdBlocker", "Settings", "Color", 0))
chkallcookies.Value = CInt(GetSetting("AdBlocker", "Settings", "All", 1))
chkusedns.Value = CInt(GetSetting("AdBlocker", "Settings", "UseDNS", 0))
chkblockallpops.Value = CInt(GetSetting("AdBlocker", "Settings", "AllPops", 1))
End Sub
Private Sub CheckDNS()
lstdns.Enabled = chkusedns.Value
End Sub
Private Sub CheckAllCookies()
lstcookies.Enabled = Not chkallcookies.Value
End Sub

Private Sub txtpopups_Change()
cmdaddpops.Enabled = Len(Trim(txtpopups.Text))
End Sub
Private Sub initFTab()
Call LoadFileLst(lstadservers, App.Path + "\data\hosts.svr")
lbladserver.Caption = Str(lstadservers.ListCount)
End Sub

Private Sub EnableApply()
cmdapply.Enabled = True
End Sub
 
Private Sub SetLstFocus(box As ListBox)
box.ListIndex = box.ListCount - 1
End Sub

Private Sub ws_Connect()
If AddStringToList(ws.RemoteHostIP + " " + Trim(txtdns.Text), frmSettings.lstdns) = True Then
Call EnableApply
Call SetLstFocus(frmSettings.lstdns)
txtdns.Text = ""
lbldns.Caption = Str(lstdns.ListCount)
End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Host not found", vbCritical, "Error"
txtdns.Text = ""
End Sub
