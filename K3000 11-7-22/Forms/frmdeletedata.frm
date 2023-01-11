VERSION 5.00
Object = "{8A0DA067-1D11-458E-9390-F81F9C64F3EB}#5.0#0"; "ProgressBar.ocx"
Begin VB.Form frmdeletedata 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DELETE RECORDS"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000D&
      Caption         =   "Check to Reset Transaction Number."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000D&
      Caption         =   "Check to Reset Count Number."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin ProgressBarPrj.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1085
      vPercView       =   -1  'True
      vForeColor      =   65280
      vBackColor      =   -2147483632
      vTextColor      =   -2147483634
      vPercCaption    =   "% Complete...."
      vUnloadProgBar  =   0
      vPercType       =   0
      vPercShape      =   0
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
      Left            =   8640
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmddelprod 
      Caption         =   "DELETE ALL DATA IN PRODUCT"
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
      Left            =   360
      TabIndex        =   3
      Top             =   4320
      Width           =   10215
   End
   Begin VB.CommandButton cmddelme 
      Caption         =   "DELETE ALL DATA IN UNIT OF MEASURE"
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
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   10215
   End
   Begin VB.CommandButton cmddelcus 
      Caption         =   "DELETE ALL DATA IN CUSTOMER"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   10215
   End
   Begin VB.CommandButton cmddelws 
      Caption         =   "DELETE ALL DATA IN WEIGHING SCALE"
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
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   10215
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "frmdeletedata.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE DATABASE RECORDS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: If you delete all record you cannot restore unless you have the back up."
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
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   8250
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
Attribute VB_Name = "frmdeletedata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim delsql As String
Dim P As Integer
Dim i As Integer
Const MaxItems As Integer = 50
Dim consec_num As String
Dim count_num As String
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelcus_Click()
If MsgBox("  Are You Sure You Want To Delete All Records?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
 Me.ProgressBar1.Visible = True
For i = 1 To MaxItems
 Sleep 100
    ProgressBar1.SetPerc i, MaxItems
  Next i
With con
.BeginTrans
delsql = "Truncate tblcustomer"
.Execute delsql
.CommitTrans
MsgBox "All Records has been deleted", vbInformation, "Thank You"
 Me.ProgressBar1.Visible = False
   Call addNewLog(currentuser, "Delete All Customer")
End With
Else
Cancel = 1
End If
End Sub

Private Sub cmddelme_Click()

If MsgBox("  Are You Sure You Want To Delete All Records?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
 Me.ProgressBar1.Visible = True
For i = 1 To MaxItems
 Sleep 100
    ProgressBar1.SetPerc i, MaxItems
  Next i
With con
.BeginTrans
'sql = "ALTER table tblunitmeasure ALTER column unitid Autoincrement(1,1)"
'delsql = "Delete * From tblunitmeasure"
delsql = "Truncate tblunitmeasure"
.Execute delsql
'.Execute sql
.CommitTrans
MsgBox "All Records has been deleted", vbInformation, "Thank You"
 Me.ProgressBar1.Visible = False
   Call addNewLog(currentuser, "Delete All Unit")
End With
Else
Cancel = 1
End If
End Sub

Private Sub cmddelprod_Click()
If MsgBox("  Are You Sure You Want To Delete All Records?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
 Me.ProgressBar1.Visible = True
For i = 1 To MaxItems
 Sleep 100
    ProgressBar1.SetPerc i, MaxItems
  Next i
With con
.BeginTrans
'sql = "ALTER table tblproduct ALTER column productid Autoincrement(1,1)"
'delsql = "Delete * From tblproduct"
delsql = "Truncate tblproduct"
.Execute delsql
'.Execute sql
.CommitTrans
MsgBox "All Records has been deleted", vbInformation, "Thank You"
 Me.ProgressBar1.Visible = False
   Call addNewLog(currentuser, "Delete All Product")
End With
Else
Cancel = 1
End If
End Sub

Private Sub cmddelws_Click()
consec_num = "00000"
count_num = "0000000"
If MsgBox("  Are You Sure You Want To Delete All Records?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
 Me.ProgressBar1.Visible = True
 For i = 1 To MaxItems
 Sleep 100
    ProgressBar1.SetPerc i, MaxItems
  Next i
With con
.BeginTrans
delsql = "Truncate tblweighing"
.Execute delsql
.CommitTrans
MsgBox "All Records has been deleted", vbInformation, "Thank You"
If Me.Check2.Value = 1 Then
     con.BeginTrans
        Call cmmd("update tblsetup set soldidcnt='" & consec_num & "'")
     con.CommitTrans
End If
If Me.Check1.Value = 1 Then
    con.BeginTrans
        Call cmmd("update tblcount set countnumber='" & count_num & "'")
    con.CommitTrans
End If
 Me.ProgressBar1.Visible = False
      Call addNewLog(currentuser, "Delete All Weighing")
End With
Else
Cancel = 1
End If
End Sub

Private Sub ProgressBar1_CurrentProgress(CurrentRecord As Integer, TotalRecords As Integer)
  If TotalRecords = 100 Then
    ProgressBar1.PERC_CAPTION = "% Complete"
  End If
End Sub

