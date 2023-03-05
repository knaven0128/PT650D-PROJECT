VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmcompany 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPANY PROFILE"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "frmcompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleMode       =   0  'User
   ScaleWidth      =   16649.78
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbunit 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "frmcompany.frx":000C
      Left            =   5280
      List            =   "frmcompany.frx":000E
      TabIndex        =   27
      Top             =   3960
      Width           =   2775
   End
   Begin VB.ComboBox cmbproduct 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "frmcompany.frx":0010
      Left            =   1800
      List            =   "frmcompany.frx":0012
      TabIndex        =   24
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtstatus 
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
      Left            =   1800
      TabIndex        =   22
      Tag             =   "txtcom"
      Text            =   "Reg"
      Top             =   5040
      Width           =   5175
   End
   Begin VB.CommandButton cmdupload 
      Caption         =   "Upload Logo"
      Height          =   255
      Left            =   11640
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   1866
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   11640
      ScaleHeight     =   1875
      ScaleWidth      =   1800
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1866
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CheckBox checkaddress 
      BackColor       =   &H8000000D&
      Caption         =   "Show on Header Layout?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   18
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CheckBox checkcontact 
      BackColor       =   &H8000000D&
      Caption         =   "Show on Header Layout?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CheckBox checkemail 
      BackColor       =   &H8000000D&
      Caption         =   "Show on Header Layout?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CheckBox checkname 
      BackColor       =   &H8000000D&
      Caption         =   "Show on Header Layout?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   1800
      Width           =   3255
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
      Left            =   1800
      TabIndex        =   14
      Tag             =   "txtcom"
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "&SET"
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
      Left            =   7080
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
   End
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
      Left            =   9000
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtaddress 
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
      Left            =   1800
      TabIndex        =   3
      Tag             =   "txtcom"
      Top             =   3120
      Width           =   6255
   End
   Begin VB.TextBox txtcontact 
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
      Left            =   1800
      TabIndex        =   2
      Tag             =   "txtcom"
      Top             =   2640
      Width           =   6255
   End
   Begin VB.TextBox txtemail 
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
      Left            =   1800
      TabIndex        =   1
      Tag             =   "txtcom"
      Top             =   2160
      Width           =   6255
   End
   Begin VB.TextBox txtcompname 
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
      Left            =   1800
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   1680
      Width           =   6255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "UNIT:"
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
      Left            =   4679
      TabIndex        =   28
      Top             =   4080
      Width           =   967
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT:"
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
      Left            =   600
      TabIndex        =   26
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT YOUR DEFAULT TO USE EVERY SCALE"
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
      Left            =   1800
      TabIndex        =   25
      Top             =   3600
      Width           =   6015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS:"
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
      Left            =   840
      TabIndex        =   23
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   7560
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SERIAL NUMBER:"
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
      TabIndex        =   13
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY PROFILE"
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
      Left            =   1920
      TabIndex        =   10
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
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   3690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   120
      Picture         =   "frmcompany.frx":0014
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ADDRESS:"
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
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT # :"
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
      Left            =   600
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS:"
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
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY NAME:"
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
      Top             =   1800
      Width           =   1695
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
Attribute VB_Name = "frmcompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numregister As String



Private Sub cmbproduct_DropDown()
Me.cmbproduct.Clear
cmbproduct.AddItem "NA"
Set rsproduct = Nothing
With rsproduct
.Open "Select * from tblproduct", con, 3, 3
Do Until .EOF
cmbproduct.AddItem !product_name
.MoveNext
Loop
End With
rsproduct.Close
End Sub

Private Sub cmbunit_DropDown()
Me.cmbunit.Clear
cmbunit.AddItem "NA"
Set rsum = Nothing
With rsum
.Open "Select * from tblunitmeasure", con, 3, 3
Do Until .EOF
cmbunit.AddItem !unit_name
.MoveNext
Loop
End With
rsum.Close
End Sub

Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub
Private Sub cmdset_Click()
numregister = "Not"
If Me.txtserial.Text = "" Then
With cmd
                .ActiveConnection = ocn
                .CommandType = adCmdText
                .CommandText = "update tblcompany " & _
                "set company_name='" & Me.txtcompname.Text & "'," & _
                "namecheck='" & Me.checkname.Value & "'," & _
                 "company_email='" & Me.txtemail.Text & "'," & _
                 "emailcheck='" & Me.checkemail.Value & "'," & _
                 "company_contact='" & Me.txtcontact.Text & "'," & _
                 "contactcheck='" & Me.checkcontact.Value & "'," & _
                 "company_address='" & Me.txtaddress.Text & "'," & _
                 "addresscheck='" & Me.checkaddress.Value & "'," & _
                 "regnum='" & Me.txtstatus.Text & "'," & _
                 "serial_number='" & Me.txtserial.Text & "'," & _
                   "de_product='" & Me.cmbproduct.Text & "'," & _
                   "de_unit='" & Me.cmbunit.Text & "'" & _
                "where CompID=" & 1
              .Execute
            End With
            MsgBox "Data totally Set!", vbInformation, "Info"
Else
 With cmd
                .ActiveConnection = ocn
                .CommandType = adCmdText
                .CommandText = "update tblcompany " & _
                "set company_name='" & Me.txtcompname.Text & "'," & _
                "namecheck='" & Me.checkname.Value & "'," & _
                "company_email='" & Me.txtemail.Text & "'," & _
                "emailcheck='" & Me.checkemail.Value & "'," & _
                "company_contact='" & Me.txtcontact.Text & "'," & _
                "contactcheck='" & Me.checkcontact.Value & "'," & _
                "company_address='" & Me.txtaddress.Text & "'," & _
                "addresscheck='" & Me.checkaddress.Value & "'," & _
                "de_product='" & Me.cmbproduct.Text & "'," & _
                "regnum='" & Me.txtstatus.Text & "'," & _
                "serial_number='" & Me.txtserial.Text & "'," & _
                "de_unit='" & Me.cmbunit.Text & "'" & _
                "where CompID=" & 1
              .Execute
    

            End With
            MsgBox "Data totally Set!", vbInformation, "Info"
End If
End Sub

Private Sub cmdupload_Click()
On Error GoTo Err:
Me.CommonDialog1.FileName = "'"
Me.CommonDialog1.Filter = "JPEG Files|*.jpg|GIF Files|*.gif|All Files*.*"
Me.CommonDialog1.ShowOpen
Me.Label8.Caption = Me.CommonDialog1.FileName
Image2.Picture = LoadPicture(Label8.Caption)
Err:
Exit Sub

End Sub

Private Sub Form_Load()
If currentposition = "Programmer" Then
MsgBox "Youre a programmer", vbInformation, "Programmer"
Else
frmcompany.txtstatus.Visible = False
frmcompany.Label9.Visible = False
frmcompany.txtcompname.Enabled = False
frmcompany.txtserial.Visible = False
frmcompany.Label6.Visible = False
End If
On Error GoTo Error:
Set rscompany = Nothing
rscompany.Open "select * from tblcompany", ocn, 3, 3
Me.txtcompname.Text = rscompany!company_name
Me.checkname.Value = rscompany!namecheck
Me.txtemail.Text = rscompany!company_email
Me.checkemail.Value = rscompany!emailcheck
Me.txtcontact.Text = rscompany!company_contact
Me.checkcontact.Value = rscompany!contactcheck
Me.txtaddress.Text = rscompany!company_address
Me.checkaddress.Value = rscompany!addresscheck
Me.txtserial.Text = rscompany!serial_number
Me.cmbproduct.Text = rscompany!de_product
Me.cmbunit.Text = rscompany!de_unit
Me.txtstatus.Text = rscompany!regnum
Error:
Exit Sub
End Sub

