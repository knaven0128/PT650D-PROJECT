VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmuseraccount 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Content"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12900
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "USER ENTRY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   13095
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
         Left            =   10920
         TabIndex        =   23
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&SAVE"
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
         Left            =   8880
         TabIndex        =   22
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox txtconfirm 
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
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   3
         Tag             =   "txtcom"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtfname 
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
         Left            =   2880
         TabIndex        =   5
         Tag             =   "txtcom"
         Top             =   2760
         Width           =   5415
      End
      Begin VB.ComboBox cmbposition 
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
         ItemData        =   "frmuseraccount.frx":0000
         Left            =   2880
         List            =   "frmuseraccount.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtuser 
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
         Left            =   2880
         TabIndex        =   1
         Tag             =   "txtcom"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtpass 
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
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "txtcom"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD:"
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
         Left            =   720
         TabIndex        =   21
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FULL NAME:"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "POSITION:"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME:"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&DELETE"
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
      Left            =   7200
      TabIndex        =   27
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&EDIT"
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
      Left            =   5280
      TabIndex        =   26
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&ADD"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   7080
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
      Left            =   9120
      TabIndex        =   24
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Click the desire user before editing!"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   4050
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "USER REGISTRATION"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   0
      Width           =   10215
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
      TabIndex        =   16
      Top             =   240
      Width           =   105
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
      TabIndex        =   15
      Top             =   960
      Width           =   3690
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME:"
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
      Left            =   8400
      TabIndex        =   12
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   360
      Picture         =   "frmuseraccount.frx":0044
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label lbluser 
      BackStyle       =   0  'Transparent
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   60
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
Attribute VB_Name = "frmuseraccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addnum As Byte
Dim admin1 As String
Dim update_data As String

Private Sub cmdadd_Click()
addnum = 1
Me.Frame1.Visible = True
Me.cmbposition.Text = "Select User..."
End Sub

Private Sub cmdcancel_Click()
With rsuser
        If MsgBox("  Are You Sure You Want To Cancel This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
         .CancelBatch
        Me.Frame1.Visible = False
        If .RecordCount > 0 Then
            .MoveFirst
             Call Emptyctl(Me, "txtcom")
             Me.Frame1.Visible = False
            Else
        End If
        End If
    End With
End Sub

Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub

Private Sub cmddelete_Click()
On Error GoTo fixDel
With rsuser
        If .RecordCount > 0 Then
            If MsgBox("  Are You Sure You Want To Delete This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
              Call addNewLog(currentuser, "Delete User" + rsuser!user_name)
            .Delete
            .UpdateBatch
            MsgBox "   Record Successfully Deleted..  ", vbExclamation, "Delete"
                Else
            .CancelUpdate
            MsgBox "  Deletion Cancelled!  ", vbOKOnly, "Cancel"
            End If
            Else
        GoTo fixDel
        End If
    End With
    Call Form_Load
fixDel:
If Err.Number = 6160 Then
End If
End Sub

Private Sub cmdedit_Click()
On Error GoTo fixEdit
Me.Frame1.Visible = True
addnum = 2
Me.txtuser.Text = rsuser!user_name
Me.txtpass.Text = rsuser!user_password
Me.cmbposition.Text = rsuser!user_position
Me.txtfname.Text = rsuser!full_name
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdsave_Click()
 If Me.cmbposition.Text = "Select User..." Then
    MsgBox "Please Select you Position!.", vbExclamation, "System Prompt"
     Me.cmbposition.SetFocus
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
If Me.txtfname.Text = "" Then
    MsgBox "Please Input Full Name!.", vbExclamation, "System Prompt"
     Me.txtfname.SetFocus
     Exit Sub
End If
If Trim$(Me.txtpass.Text) = Trim$(Me.txtconfirm.Text) Then
Else
MsgBox "Please Confirm password!", vbCritical
Me.txtconfirm.Text = ""
Me.txtconfirm.SetFocus
Exit Sub
End If
If currentposition = "Administrator" And Me.cmbposition.Text = "Programmer" Then
MsgBox "You cannot add or edit a programmer!"
ElseIf currentposition = "Weigher" And Me.cmbposition.Text = "Programmer" Then
MsgBox "You cannot add or edit a programmer!"
Else
update_data = dataUpdated()
Select Case addnum
 Case 1
rsuser.MoveFirst
While rsuser.EOF = False
If Trim$(Me.txtuser.Text) = rsuser!user_name Then
MsgBox "User Name Already Exist!", vbOKOnly + vbCritical
Me.txtuser.SetFocus
Exit Sub
End If
rsuser.MoveNext
Wend

              With cmd
                .ActiveConnection = ocn
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tbluser " & _
                "(user_name,user_password,user_position,full_name,date_added)" & _
                " VALUES(" & _
                 "'" & Me.txtuser.Text & "'," & _
                "'" & Me.txtpass.Text & "'," & _
                "'" & Me.cmbposition.Text & "'," & _
                "'" & Me.txtfname.Text & "'," & _
                "'" & sDateString & "'" & _
                ")"
                  .Execute
                Call addNewLog(currentuser, "Add User - username name: " + Me.txtuser.Text + " - user position:" + Me.cmbposition.Text)
            End With
            
            
            Call Emptyctl(Me, "txtcom")
       Call Form_Load
           
            MsgBox "   Record Successfully Add to the System.   ", vbOKOnly, "Success!"
Case 2
 With cmd
                .ActiveConnection = ocn
                .CommandType = adCmdText
                .CommandText = "update tbluser " & _
                "set user_name='" & Me.txtuser.Text & "'," & _
                 "user_password='" & Me.txtpass.Text & "'," & _
                 "user_position='" & Me.cmbposition.Text & "'," & _
                 "full_name='" & Me.txtfname.Text & "'" & _
                " where userid=" & rsuser!UserID
              .Execute
              Call addNewLog(currentuser, "Update User - " + "ID: " + CStr(rsuser!UserID) + update_data)
            End With
            Call Emptyctl(Me, "txtcom")
                Call Form_Load
               
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
End Select
    Me.Frame1.Visible = False
End If
End Sub



Private Sub Form_Load()
sDateString = Format(Now, "yyyy/mm/dd")
Me.Timer1.Enabled = True
Me.Timer1.Interval = 1000
admin1 = "i"
Call dbconnect

End Sub


Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub

Private Sub dbconnect()
If currentposition = "Programmer" Then
Set rsuser = Nothing
rsuser.Open "select * from tbluser", ocn, 3, 3
Set DataGrid1.DataSource = rsuser
Call datasize
ElseIf currentposition = "Seller" Then
Set rsuser = Nothing
rsuser.Open "select * from tbluser where user_position= 'Administrator' or user_position= 'Weigher' or user_position= 'Seller'", ocn, 3, 3
Set DataGrid1.DataSource = rsuser
 Call datasize
ElseIf currentposition = "Administrator" Then
Set rsuser = Nothing
rsuser.Open "select * from tbluser where user_position= 'Administrator' or user_position= 'Weigher'", ocn, 3, 3
Set DataGrid1.DataSource = rsuser
 Call datasize
End If
End Sub
Private Sub datasize()
With DataGrid1
If currentposition = "Weigher" Then
      .Columns.Item(0).Visible = False
End If
      
            .Columns.Item(1).width = 1500
            .Columns.Item(2).width = 1500
            .Columns.Item(3).width = 2000
            .Columns.Item(4).width = 2500
            .Columns.Item(5).width = 1500
End With
End Sub

Private Function dataUpdated() As String
If rsuser.RecordCount > 0 Then
If rsuser!user_name <> Me.txtuser.Text Then
dataUpdated = dataUpdated + " - user name: " + rsuser!user_name
End If
If rsuser!user_password <> Me.txtpass.Text Then
dataUpdated = dataUpdated + " - password: " + "password updated"
End If

If rsuser!user_position <> Me.cmbposition.Text Then
dataUpdated = dataUpdated + " - user position: " + rsuser!user_position
End If

If rsuser!full_name <> Me.txtfname.Text Then
dataUpdated = dataUpdated + " - user fullname: " + rsuser!full_name
End If
End If
End Function
