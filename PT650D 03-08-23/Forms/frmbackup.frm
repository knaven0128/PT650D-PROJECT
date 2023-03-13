VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmbackup 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Database"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbackupsql 
      Caption         =   "&Back Up"
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
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdrestore 
      Caption         =   "&Restore"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   7815
      TabIndex        =   2
      Top             =   120
      Width           =   7815
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   720
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdbackup 
      Caption         =   "&Back Up"
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
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6240
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   195
      Left            =   9000
      Top             =   1080
      Width           =   1020
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackup_Click()
Dim FilePath As String
On Error GoTo Err:
'Now let’s do the coding
Me.CommonDialog1.CancelError = True
              '  I let the name of the controls to its default name, on button backup put this line of codes
       Me.CommonDialog1.Filter = "*.mdb"
       Me.CommonDialog1.FileName = "weighingdb"
'This line code just declare that you can only save the file as .mdb
      Me.CommonDialog1.ShowSave
      FilePath = Replace(CommonDialog1.FileName, "\" & CommonDialog1.FileTitle, "")
      MsgBox FilePath
      con.Close
'IT shows the dialog box to located where do you want to save the file and what would be the name of the file to be backup
       If Me.CommonDialog1.FileName <> "" Then
'Checks whether open button is clicked
      
        FileCopy App.Path & "\database\weighingdb.mdb", Me.CommonDialog1.FileName & ".mdb"
        
       'It just inform the user that the file is saved on  that specific location
         Call progress
       con.Open
       End If
Err:
       Timer1.Enabled = True
       Exit Sub
End Sub


 
Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub






Private Sub cmdRestore_Click()
' Click it to check mark and click OK
If MsgBox("You want to restore database?" & vbNewLine & "All current data will be replace!", vbQuestion + vbYesNo, "Prompt") = vbYes Then
Dim FSO As New FileSystemObject
Dim Src As String
Dim Dest As String
Dim xten As String
con.Close
Me.CommonDialog2.Filter = "*.mdb" 'to filter the file
Me.CommonDialog2.ShowOpen 'to show where you want to fetch/get your backup
xten = Me.CommonDialog2.FileTitle
Src = Me.CommonDialog2.FileName 'source
Dest = App.Path & "\database\weighingdb.mdb" 'the database location


With FSO

If xten = "weighingdb.mdb" Then
.DeleteFile Dest
 .CopyFile Src, Dest, True
 Call progress1
Else
MsgBox "Cannot Restore Database!" & vbnextline & "Please Contact the Programmer"
con.Open
End If
End With
Else
Cancel = 1
End If
End Sub

Private Sub Command1_Click()
 Open App.Path & "\database\run.bat" For Output As #1
 Print #1, "call set path=C:\Program Files (x86)\MySQL\MySQL Server 5.0\bin"
 Print #1, "call mysqldump -uroot -p12345  weighingscaledb > c:\weighingscaledb.sql"
Close #1
Shell App.Path & "\database\run.bat", vbHide

End Sub

Private Sub Command2_Click()
Dim cmd As String
Dim cmd1 As String
    Screen.MousePointer = vbHourglass
    DoEvents
    cmd = "set path=C:\Program Files (x86)\MySQL\MySQL Server 5.0\bin"
    cmd1 = "mysqldump -uroot -p12345  weighingscaledb > c:\weighingscaledb.sql"
    Call execCommand(cmd)
    Call execCommand(cmd1)
    Screen.MousePointer = vbDefault
    MsgBox "done"
End Sub

Private Sub Timer1_Timer()
     Static i As Integer
    
    i = i + 1
    DrawProgress Picture1, i
    
    If i = 100 Then
    MsgBox "Database backup complete!" & Me.CommonDialog1.FileName & ".mdb"
    Timer1.Enabled = False
    Me.cmdbackup.Enabled = True
    Me.cmdclose.Enabled = True
    Me.cmdrestore.Enabled = True
    Me.Label1.Caption = "Transaction Done...."
    i = 0
End If
End Sub
Private Sub progress()
Dim x As Integer
Dim y As Integer
Me.cmdrestore.Enabled = False
Me.cmdbackup.Enabled = False
Me.cmdclose.Enabled = False

 Me.Label1.Caption = "Please Wait...."
    Timer1.Interval = 100
    Timer1.Enabled = True
    x = Screen.TwipsPerPixelX
    y = Screen.TwipsPerPixelY
    
    With Picture1
        Shape1.Move .Left - (x), .Top - (y), .width + (3 * x), .height + (3 * y)
    End With
    
    DrawProgress Picture1, 0, 0, 100, vbWhite, vbBlack
End Sub

Private Sub progress1()
Dim x As Integer
Dim y As Integer
Me.cmdbackup.Enabled = False
Me.cmdrestore.Enabled = False
Me.cmdclose.Enabled = False

 Me.Label1.Caption = "Please Wait...."
    Timer2.Interval = 100
    Timer2.Enabled = True
    x = Screen.TwipsPerPixelX
    y = Screen.TwipsPerPixelY
    
    With Picture1
        Shape1.Move .Left - (x), .Top - (y), .width + (3 * x), .height + (3 * y)
    End With
    
    DrawProgress Picture1, 0, 0, 100, vbWhite, vbBlack
End Sub


Private Sub Timer2_Timer()
Static i As Integer
    
    i = i + 1
    DrawProgress Picture1, i
    
    If i = 100 Then
    MsgBox "Restore Database Complete!" & vbNewLine & "System will be shutdown!", vbOKOnly
    End
    Timer2.Enabled = False
    Me.cmdbackup.Enabled = True
    Me.cmdclose.Enabled = True
     Me.cmdrestore.Enabled = True
    Me.Label1.Caption = "Transaction Done...."
    i = 0
End If
End Sub

