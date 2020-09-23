VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "PyroNet Explorer"
   ClientHeight    =   7395
   ClientLeft      =   -90
   ClientTop       =   465
   ClientWidth     =   10755
   Icon            =   "browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9360
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   480
      Picture         =   "browser.frx":0442
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   255
      Left            =   14520
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   15135
      ExtentX         =   26696
      ExtentY         =   16325
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "about:blank"
      Top             =   600
      Width           =   14055
   End
   Begin VB.Label Label2 
      Caption         =   "Made by: SpitFire"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "PyroNet Explorer -=- Ready"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   10320
      Width           =   11055
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AllowPopup As Boolean 'This is for Pop-up windows
Private Sub Close_Click()
End
End Sub

Private Sub cmdBack_Click()
'Go back one page
WebBrowser1.GoBack
End Sub

Private Sub cmdForward_Click()
'go forward one page
WebBrowser1.GoForward
End Sub

Private Sub cmdGo_Click()
'Go to web page
WebBrowser1.Navigate txtAddress.Text
lblStatus.Caption = "Going to: " & txtAddress.Text
End Sub

Private Sub cmdRefresh_Click()
'Refresh page
WebBrowser1.Refresh
End Sub

Private Sub cmdStop_Click()
'Stop loading
WebBrowser1.Stop
End Sub

Private Sub mnuAbout_Click()

End Sub

Private Sub mnuexit_Click()
'Exit program
Unload Me
End Sub

Private Sub mnuOptionsAllow_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
WebBrowser1.GoHome
End Sub

Private Sub Command2_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Form_Load()
WebBrowser1.GoHome
Form1.WindowState = 2 - Maximized
If Form1.WindowState = 0 - Normal Then Form1.WindowState = 2 - Maximized
End Sub

Private Sub List1_Click()

End Sub

Private Sub Timer1_Timer()
If Form1.WindowState = 0 - Normal Then Form1.WindowState = 2 - Maximized
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
lblStatus.Caption = "Done Loading"
Form1.Caption = "PyroNet Explorer -=- " & WebBrowser1.LocationName
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_DownloadBegin()
'Starting download
lblStatus.Caption = "Loading..."
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_DownloadComplete()
'Done downloading
lblStatus.Caption = "Download Done!"
txtAddress.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
lblStatus.Caption = "Done Loading!"
Form1.Caption = "PyroNet Explorer -=- " & WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
'This will allow a pop-up window to load or to be blocked!
If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub

Function FileExist(vFile As String) As Boolean
    On Error Resume Next
    FileExist = False
    If Dir$(vFile) <> "" Then: FileExist = True
End Function
