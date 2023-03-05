VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8835
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   1920
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&EXIT"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&LOGIN"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Show Password?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      MaskColor       =   &H00C00000&
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "txtcom"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Tag             =   "txtcom"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmLogin.frx":4888A
      Left            =   2520
      List            =   "frmLogin.frx":4889D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8160
      TabIndex        =   12
      Top             =   3720
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Attempts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   11
      Top             =   3720
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN SESSION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   10215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Please fill all the fields needed!"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   360
      Picture         =   "frmLogin.frx":488DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Index           =   1
      Left            =   -720
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FS As New FileSystemObject
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim A As TextStream

 
Dim sTimeString As String
Dim attempt As Integer
Dim cmd As String


Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
Me.txtpass.PasswordChar = ""
Else
Me.txtpass.PasswordChar = "*"
End If
End Sub

Private Sub cmdlogin_Click()
 If Me.Combo1.Text = "Select User..." Then
    MsgBox "Please Select you Position!.", vbExclamation, "System Prompt"
     Me.Combo1.SetFocus
    Exit Sub
       End If
If Me.txtuser.Text = "" Then
    MsgBox "Don't Leave User Name Empty.", vbExclamation, "System Prompt"
     Me.txtuser.SetFocus
    Exit Sub
       End If
If Me.txtpass.Text = "" Then
    MsgBox "Don't Leave Password Empty.", vbExclamation, "System Prompt"
     Me.txtpass.SetFocus
    Exit Sub
       End If
       Set rsuser = Nothing
   With rsuser
    .Open "select * from tbluser where user_Position like '" & Me.Combo1.Text & "' and User_Name like '" & Me.txtuser.Text & "' and  user_Password like '" & txtpass.Text & "'", ocn, 3, 3
        If rsuser.EOF Then
            MsgBox "Login Error Please Check Data!", vbCritical, "Please Try Again"
            attempt = attempt + 1
            Me.Label6.Caption = attempt
            If attempt = 3 Then
              MsgBox "Error: Attempt Limit exceed" & vbNewLine & "This program will now end!", vbExclamation, "Login Error"
              End
              rsuser.Close
            End If
            Me.txtuser.Text = ""
            Me.txtpass.Text = ""
            Me.txtuser.SetFocus
           
        ElseIf MsgBox("Attempting to login", vbYesNo + vbQuestion, "Successfully") = vbYes Then
        Cancel = 1
            currentname = rsuser!full_name
            currentposition = CStr(Me.Combo1.Text)
            currentuser = CStr(Me.txtuser.Text)
                  Call addNewLog(CStr(Me.txtuser.Text), "User Login")
            Me.txtuser.Text = ""
            Me.txtpass.Text = ""
            Unload Me
            MainForm.Show 1
      
              End If
 
 End With

End Sub

Private Sub cmdExit_Click()
If MsgBox("End this application?", vbYesNo + vbQuestion, "Shutdown the Sytem") = vbYes Then End
End Sub

Private Sub Combo2_DropDown()
Combo2.Clear
With rstrucking
.Open "Select * from tbltruck", con, 1, 2
Do Until .EOF
Combo2.AddItem !truckname
.MoveNext

Loop
rstrucking.Close
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 Call cmdlogin_Click
 End If
End Sub

Private Sub Form_Load()
 servno = "localhost"
    userr = "root"
    passs = "12345"
If openn = True Then
End If
'Label7.Caption = Label7.Caption & Space(150)
Me.Combo1.Text = "Select User..."
sDateString = Format(Date, "m/dd/yyyy")
sTimeString = Format(Time, "hh:mm:ss AM/PM")
End Sub




Private Sub Timer1_Timer()
'Dim str As String
'str = frmLogin.Label7.Caption
'str = Mid$(str, 2, Len(str)) + Left(str, 1)
'frmLogin.Label7.Caption = str

End Sub
