VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Mini-Filemanager by JBK"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12465
   Icon            =   "Newmfform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14.182
   ScaleMode       =   7  'Zentimeter
   ScaleWidth      =   21.987
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7440
      Top             =   4320
   End
   Begin VB.PictureBox picUsage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   28
      Top             =   480
      Width           =   990
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8400
      ScaleHeight     =   100
      ScaleMode       =   0  'Benutzer
      ScaleWidth      =   100
      TabIndex        =   27
      Top             =   480
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   480
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "End"
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   7440
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   7200
      Top             =   5640
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Source:"
      Height          =   6015
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Network"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "Newmfform.frx":27A2
         Style           =   1  'Grafisch
         TabIndex        =   23
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   18
         Top             =   5520
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   3990
         Left            =   2160
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   16
         Top             =   5040
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Height          =   4140
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   7
         Top             =   6000
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Destination:"
      Height          =   6015
      Left            =   4440
      TabIndex        =   14
      Top             =   960
      Width           =   2535
      Begin VB.CommandButton Command4 
         Caption         =   "Network"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Picture         =   "Newmfform.frx":38DC
         Style           =   1  'Grafisch
         TabIndex        =   24
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   10
         Top             =   6000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.DirListBox Dir2 
         Height          =   5040
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   4200
      Picture         =   "Newmfform.frx":4A16
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Refresh all drives and folders"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Create new folder"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Move selected file"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete selected folder"
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete selected file"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Copy selected file"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   615
      Left            =   3240
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mini-Filemanager"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnosf 
         Caption         =   "&Open selected file"
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
      End
      Begin VB.Menu mnend 
         Caption         =   "&End"
      End
   End
   Begin VB.Menu mnuWM 
      Caption         =   "&Windows-Manager"
      Begin VB.Menu mnuss 
         Caption         =   "&Shutdown system"
      End
      Begin VB.Menu mnlo 
         Caption         =   "&Logoff"
      End
      Begin VB.Menu mnur 
         Caption         =   "&Reboot"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QueryObject As Object

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
    Dim s As Integer
    Dim dta As String




Private Sub Command10_Click()
FrmLittleBarGraph.Show
    '   ^ open cd rom tray
End Sub


Private Sub Command11_Click()
    mciSendString "set CDAudio door closed", t, 127, 0
    '   ^ close cd rom tray
Command10.Visible = True
Command11.Visible = False
End Sub

Private Sub Command4_Click()
On Error GoTo error
Dim net
net = InputBox("Insert the path of the network drive! Example: \\pc2\c", "Network", "\\pc2\d")
Dir2.Path = net
error:
If Err.Number = 68 Then
MsgBox "No disk in Drive!", vbCritical, "Problem..."
Drive1.Drive = "C:\"
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

Exit Sub

'On Error GoTo error

'error:
'MsgBox "No network available!"
'Exit Sub
End If

End Sub

Private Sub Command2_Click()
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

End Sub

Private Sub Command3_Click()
MsgBox "Are you sure you want to delete this file?", vbYesNo, "Delete..."
If vbYes Then
kill
Else
Exit Sub
End If
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

End Sub

Private Sub Command5_Click()
Dim a
a = InputBox("Do you want to delete the folder: " + Form1.Text2.Text + "  ? Type in YES to delete or NO to cancel.", "Delete folder.", "yes")

If a = "yes" Then
rmv
End If
If a = "No" Then
Exit Sub
End If
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh
End Sub

Private Sub Command6_Click()
filecopy3
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh
End Sub

Private Sub Command7_Click()
filecopy4
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

End Sub

Private Sub Command8_Click()
make2
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

End Sub

Private Sub Command9_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label6.Caption = File1.ListCount
Text1.Text = "The folder: " + Dir1.Path + " includes: " + Label6.Caption + " File(s)."
Text3.Text = Dir1.Path
End Sub

Private Sub Dir2_Change()
Text2.Text = Dir2.Path
Form1.Caption = "Mini-Filemanager by JBK" + "   -   " + Dir2.Path

End Sub

Private Sub Drive1_Change()
On Error GoTo error
Dir1.Path = Drive1.Drive
error:
If Err.Number = 68 Then
MsgBox "No disk in Drive!", vbCritical, "Problem..."
Drive1.Drive = "C:\"
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

Exit Sub

End If
End Sub


Private Sub Drive2_Change()
On Error GoTo error

Dir2.Path = Drive2.Drive
error:
If Err.Number = 68 Then
MsgBox "No disk in Drive!", vbCritical, "Problem..."
Drive1.Drive = "C:\"
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

Exit Sub
End If

End Sub

Private Sub File1_Click()
On Error GoTo error
Text3.Text = File1.Path
Dim lange
lange = FileSystem.FileLen(File1.Path + "\" + File1.filename)
Label8.Caption = lange / 1000000
Label9.Caption = "The file: " + File1.filename + " is: " + Label8.Caption + " Megabytes."

Text4.Text = "The file: " + File1.filename + " is: " + Label8.Caption + " Megabytes."


error:
Exit Sub
End Sub

Private Sub File1_DblClick()
'Shell File1.Path + "\" + File1.filename, vbNormalFocus
Print fHandleFile(Dir1.Path + "\" + File1.filename, WIN_NORMAL)
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Shell File1.Path + "\" + File1.filename, vbNormalFocus
End If
End Sub

Private Sub Form_Load()
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize

File1.Path = Dir1.Path
Label6.Caption = File1.ListCount
Text1.Text = "The folder: " + Dir1.Path + " includes: " + Label6.Caption + " File(s)."
Text2.Text = Dir2.Path
Form1.Caption = "Mini-Filemanager by JBK" + "   -   " + Dir2.Path
Text3.Text = Dir1.Path
'// Centering the form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    '// Change Label and Timer prefs
    Label1.Caption = ""
    Timer3.Enabled = True
    Timer3.Interval = 100
    Label1.Width = 4335
    Label1.Font = "Courier New"
    Label1.Font.Size = 8
    Label1.Font.Bold = True
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Label1_Click()
Print fHandleFile("http://jbk_prog.tripod.com/programms", WIN_NORMAL)
End Sub

Private Sub mnend_Click()
End
End Sub

Private Sub mnlo_Click()
Dim dummy
Dim answer
answer = MsgBox("Do you want Mini-Filemanager to Logoff your user account?", vbYesNo, "System Shutdown...")
If answer = vbYes Then
dummy = ExitWindowsEx(EWX_FORCE, 4)
dummy = ExitWindowsEx(EWX_LOGOFF, 0)
End If

End Sub

Private Sub mnosf_Click()
Shell File1.Path + "\" + File1.filename, vbNormalFocus

End Sub

Private Sub mnur_Click()
Dim dummy
Dim answer
answer = MsgBox("Do you want Mini-Filemanager to Reboot your system?", vbYesNo, "System Shutdown...")
If answer = vbYes Then
dummy = ExitWindowsEx(EWX_FORCE, 4)
dummy = ExitWindowsEx(EWX_REBOOT, 2)
End If

End Sub

Private Sub mnuss_Click()
Dim dummy
Dim answer
answer = MsgBox("Do you want Mini-Filemanager to Shutdown your system?", vbYesNo, "System Shutdown...")
If answer = vbYes Then
dummy = ExitWindowsEx(EWX_FORCE, 4)
dummy = ExitWindowsEx(EWX_SHUTDOWN, 1)
End If

End Sub

Private Sub Timer2_Timer()
   Dim Ret As Long
    Dim Which As Long
    'query the CPU usage
    Ret = QueryObject.Query
    If Ret = -1 Then
        Timer1.Enabled = False
        lblCpuUsage.Caption = ":-("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) & "%"
        DoEvents
    DoEvents
    End If



End Sub


Private Sub Command1_Click()
On Error GoTo error
Dim net
net = InputBox("Insert the path of the network drive! Example: \\pc2\c", "Network", "\\pc2\c")
Dir1.Path = net

error:
If Err.Number = 68 Then
MsgBox "No disk in Drive!", vbCritical, "Problem..."
Drive1.Drive = "C:\"
Dir1.Refresh
Dir2.Refresh
Drive1.Refresh
Drive2.Refresh
File1.Refresh

Exit Sub

'On Error GoTo error

'error:
'MsgBox "No network available!"
'Exit Sub
End If
End Sub



Private Sub Text5_Change()
End Sub

Private Sub Timer3_Timer()
    '// put your text here in dta string
    dta = "JBK = Dimitrios Badanis-Kapirnas. JBK-Freeware was here! Visit: http://jbk_prog.tripod.com/programms" & SPACE$(70)
    s = s + 1
    Label1.Caption = Mid(dta, 1, s)
    If Len(Label1.Caption) >= 71 Then Label1.Caption = Right(Label1.Caption, 70)


    If s = Len(dta) Then
        Label1.Caption = ""
        s = 0
    End If
End Sub

