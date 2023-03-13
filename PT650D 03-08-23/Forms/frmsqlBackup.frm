VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsqlBackup 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Back Up Database"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdclose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtDestination 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   6315
   End
   Begin VB.CommandButton cmdexport 
      Caption         =   "&BACKUP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&BROWSE...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInform 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait the data is exporting......."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1080
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "(0%)..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FILE DESTINATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   120
      Picture         =   "frmsqlBackup.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "BACK UP DATABASE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   1920
      TabIndex        =   10
      Top             =   0
      Width           =   10215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click the browse button to browse your file destination."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   5550
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click the backup button to backup data."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   1080
      Width           =   3990
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmsqlBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vdate As String
Dim mintCount As Integer, mintPause As Integer
Dim strDate As String
Dim fromdate As String
Dim todate As String
Dim filePath As String
Dim sqlPath As String
Dim qoute As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdbrowse_Click()
On Error GoTo Err:
'Now let’s do the coding
Me.CommonDialog1.CancelError = True
              '  I let the name of the controls to its default name, on button backup put this line of codes
       Me.CommonDialog1.Filter = "*.sql"
       Me.CommonDialog1.FileName = "weighingscaledb"
'This line code just declare that you can only save the file as .mdb
      Me.CommonDialog1.ShowSave
      Me.txtDestination.Text = Replace(Me.CommonDialog1.FileName, "weighingscaledb", "")
      
'IT shows the dialog box to located where do you want to save the file and what would be the name of the file to be backup

Err:
       Exit Sub
End Sub

Private Sub runBat()

End Sub

Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub

Private Sub cmdexport_Click()
If Me.txtDestination.Text = "" Then
MsgBox "Please select the file destination first!" & vbNewLine & "Click the browse button."
Else
'If InStr(Me.txtDestination.Text, " ") > 0 Then
'MsgBox "Please select different destination.", vbCritical
'Exit Sub
'End If
filePath = App.Path & "\database\run.bat"

       If Me.CommonDialog1.FileName <> "" Then
'Checks whether open button is clicked
      
        Open App.Path & "\database\run.bat" For Output As #1
        Print #1, "call set path=C:\Program Files (x86)\MySQL\MySQL Server 5.0\bin"
        Print #1, "call mysqldump -uroot -p12345  --add-drop-database --databases weighingscaledb >  " & Chr$(34) & CommonDialog1.FileName & " .sql " & Chr$(34)
        Close #1
        Shell Chr$(34) & filePath & Chr$(34), vbNormalFocus
        
        
        
        
       'It just inform the user that the file is saved on  that specific location
         Timer1.Enabled = True
         Me.cmdbrowse.Enabled = False
        Me.cmdclose.Enabled = False
        Me.cmdexport.Enabled = False
          Call addNewLog(currentuser, "Back up All Records")
       End If
End If
End Sub

Private Sub Timer1_Timer()
 Call CountMe
    lblcount.Visible = True
    lblInform.Visible = True
'    lblCBK.Visible = True
    progStat.Visible = True
    progStat.Value = progStat.Value + 2
   
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
If progStat.Value = 100 Then
   If MsgBox("Backup data is Succesful!Open file destination?", vbYesNo + vbQuestion, "Successfully") = vbYes Then
        Cancel = 1
        ShellExecute 0, vbNullString, Me.txtDestination.Text, vbNullString, vbNullString, 1

    End If
    progStat.Value = 0
    mintCount = 0
    Me.cmdbrowse.Enabled = True
    Me.cmdclose.Enabled = True
    Me.cmdexport.Enabled = True
    Timer1.Enabled = False
    lblcount.Visible = False
    lblInform.Visible = False
   progStat.Visible = False
   Me.txtDestination.Text = ""
Else
    If txtDestination.Text = "" Then
     progStat.Value = 0
     
       'Your function, can be anything. Open another form, frmMain.show... Ect.
    End If
    End If
End Sub
Private Sub CountMe()
   mintPause = mintPause + 1
   
    If mintCount < 0 Then
        mintCount = mintCount + 1
        lblcount.Caption = "(" & mintCount & "%)..."
         
    ElseIf mintCount < 100 Then
        mintCount = mintCount + 2
        lblcount.Caption = "(" & mintCount & "%)..."
        
    End If
    
    If mintPause = 100 Then
        lblcount.Caption = "(0%)..."
        lblInform.Caption = "Please wait the data is exporting......"
    ElseIf mintPause > 180 Then
   End If
End Sub


