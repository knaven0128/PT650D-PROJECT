VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogs 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYSTEM LOGS"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8160
      TabIndex        =   8
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "CL&EAR LOGS"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   480
   End
   Begin VB.TextBox txtsearch 
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
      Left            =   1560
      TabIndex        =   1
      Tag             =   "txtcom"
      Top             =   2400
      Width           =   8415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8493
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
      Caption         =   "Note: After clearing logs there's no other way to recover data logs!"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   6645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH LOGS:"
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
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
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
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
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
      Left            =   6600
      TabIndex        =   4
      Top             =   1800
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
      Left            =   5880
      TabIndex        =   3
      Top             =   2040
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
      Left            =   6600
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM LOGS"
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
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   240
      Picture         =   "frmLogs.frx":0000
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
      Left            =   -360
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
If MsgBox("  Are You Sure You Want To Clear All Logs?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
With con
.BeginTrans
delsql = "Truncate tbllogs"
.Execute delsql
.CommitTrans
Call addNewLog(currentuser, "Clear all Logs")
Call dbconnect
MsgBox "All Logs has been cleared", vbInformation, "Thank You"
End With
Else
Cancel = 1
End If
End Sub


Private Sub Form_Load()
Timer1.Interval = 1000
Call dbconnect

End Sub

Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub
Private Sub dbconnect()
Set rslogs = Nothing
rslogs.Open "select * from tbllogs ", con, 3, 3
Set DataGrid1.DataSource = rslogs
    Call datasize
End Sub
Private Sub datasize()
With DataGrid1
.WrapCellPointer = True
            .Columns.Item(0).Visible = False
            .Columns.Item(1).width = 2500
            .Columns.Item(2).width = 7500
            .RowHeight = 500
              .Columns.Item(3).width = 2500
End With
End Sub

Private Sub txtsearch_Change()
If Trim$(Me.txtsearch.Text) = "" Then
    Set rslogs = Nothing
    rslogs.Open "select * from tbllogs order by logsid ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rslogs
    Call datasize
      Set rslogs = Nothing
    rslogs.Open "select * from tbllogs order by logsid ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.txtsearch.Text) <> "" Then
Set rslogs = Nothing
    rslogs.Open "select * from tbllogs where username like'%" & Me.txtsearch.Text & "%'  OR datelogs like'%" & Me.txtsearch.Text & "%'  OR actionlog like'%" & Me.txtsearch.Text & "%'", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rslogs
    Call datasize
End If
End Sub
