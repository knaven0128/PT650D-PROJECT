VERSION 5.00
Begin VB.Form frmregister 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   3240
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8385
   Icon            =   "frmregister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdcancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdok 
         Appearance      =   0  'Flat
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtpassword 
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
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   6
         Tag             =   "txtcom"
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdset 
      Appearance      =   0  'Flat
      Caption         =   "&Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtserial 
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
      Left            =   120
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SERIAL NUMBER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   10215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weighing Scale System "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   4695
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
   Begin VB.Menu mgsn 
      Caption         =   "Help"
      Begin VB.Menu msn 
         Caption         =   "Generate SN"
      End
      Begin VB.Menu mabout 
         Caption         =   "About"
      End
      Begin VB.Menu mcu 
         Caption         =   "Contact Us"
      End
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTimeString As String
Dim myid As Integer
Private Sub cmdcancel_Click()
Me.Frame1.Visible = False
End Sub

Private Sub cmdok_Click()
If Trim$(Me.txtpassword.Text) = "Babyrr0403" Then
Me.txtpassword.Text = ""
Unload Me
frmLogin.Show 1
Else
MsgBox "Error: Password Incorrect", vbCritical
Me.txtpassword.Text = ""
Me.Frame1.Visible = False
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 Call cmdok_Click
 End If
End Sub
Private Sub cmdset_Click()
If Me.txtserial.Text = "0" Then
MsgBox "Invalid Serial Number!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Me.txtserial.Text = ""
ElseIf Trim$(Me.txtserial.Text) = serial Then
ocn.BeginTrans
Call cmmd1("update tblcompany set regnum='" & Cnumber & "'")
ocn.CommitTrans
MsgBox "Serial Numeber Succesfully Registered!" & vbNewLine & "Welcome and Enjoy", vbInformation, "Succesfully"
Unload Me
frmLogin.Show 1
Else
MsgBox "Invalid Serial Number!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Me.txtserial.Text = ""

End If
End Sub


Private Sub cmdExit_Click()
If MsgBox("End this application?", vbYesNo + vbQuestion, "Shutdown the Sytem") = vbYes Then End
End Sub

Private Sub Form_Load()
If openndb = True Then
End If
Set rscompany = Nothing
rscompany.Open "select * from tblcompany", ocn, adOpenStatic, adLockOptimistic, adCmdText
serial = rscompany!serial_number
register = rscompany!regnum
Cnumber = "Reg"
If register = "Not" Then
MsgBox "This Program is not registered!"
Else
rscompany.Close
Unload Me
frmLogin.Show
End If
sDateString = Format(Date, "m/dd/yyyy")
sTimeString = Format(Time, "hh:mm:ss AM/PM")
End Sub

Private Sub mabout_Click()
  MsgBox "Info: " & "Weighing Scale System!" & vbNewLine & "Created by Knaven Rey Sarroza.", vbOKOnly + vbInformation
  
End Sub

Private Sub mcu_Click()
  MsgBox "Contact Us " & vbNewLine & "Mobile Number: 09277469736" & vbNewLine & "Email: sarrozaknavenrey28@gmail.com", vbOKOnly + vbInformation
End Sub

Private Sub msn_Click()
Frame1.Visible = True
Me.txtpassword.SetFocus
End Sub
