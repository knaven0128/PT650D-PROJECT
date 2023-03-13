VERSION 5.00
Begin VB.Form frmscaleoffline 
   BackColor       =   &H8000000D&
   Caption         =   "SCALE OFFLINE"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CA&NCEL"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
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
      Left            =   1680
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   2040
      Width           =   3615
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
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "txtcom"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "SCALE OFFLINE"
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
      TabIndex        =   2
      Top             =   0
      Width           =   10215
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
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1230
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
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   360
      Picture         =   "frmscaleoffline.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
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
      TabIndex        =   3
      Top             =   1080
      Width           =   3930
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmscaleoffline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim userp As String

Private Sub cmdcancel_Click()
            If frmWT.ToggleButton1.Value = True Then
            frmWT.ToggleButton1.Value = False
            Unload Me
            ElseIf frmWT.ToggleButton1.Value = False Then
            frmWT.ToggleButton1.Value = True
            Unload Me
            End If
End Sub

Private Sub cmdok_Click()
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
          Set rsscaleuser = Nothing
   With rsscaleuser
    .Open "select  * from tbluser where user_Position= '" & userp & "' and User_Name='" & Me.txtuser.Text & "' and  user_Password='" & txtpass.Text & "'", ocn, adOpenDynamic, adLockOptimistic
        If rsscaleuser.EOF Then
            MsgBox "Login Error Please Login as Administrator or Check Data!", vbCritical, "Please Try Again"
            Me.txtuser.Text = ""
            Me.txtpass.Text = ""
        ElseIf MsgBox("Access Granted!", vbOKCancel + vbQuestion, "WELCOME USER") = vbOK Then
            Cancel = 1
            Call addNewLog(currentuser, "Scale Offline - granted by: " + txtuser.Text)
            Me.txtuser.Text = ""
            Me.txtpass.Text = ""
            frmWT.txtkilo.Locked = False
            frmWT.Timer1.Enabled = False
            Unload Me
              End If
 End With

End Sub

Private Sub Form_Load()
Set rsscaleuser = Nothing
userp = "Administrator"
End Sub

