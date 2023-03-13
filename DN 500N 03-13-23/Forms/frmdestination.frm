VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdestination 
   BackColor       =   &H8000000D&
   Caption         =   "Destination"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "DESTINATION"
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
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtdestination 
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
         TabIndex        =   3
         Tag             =   "txtcom"
         Top             =   600
         Width           =   3615
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
         Left            =   6120
         TabIndex        =   2
         Top             =   4320
         Width           =   1815
      End
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
         Left            =   8040
         TabIndex        =   1
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION NAME:"
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
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
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
      Left            =   2040
      TabIndex        =   9
      Tag             =   "txtcom"
      Top             =   2280
      Width           =   5175
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
      Left            =   7800
      TabIndex        =   8
      Top             =   6720
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
      Left            =   2040
      TabIndex        =   7
      Top             =   6720
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
      Left            =   3960
      TabIndex        =   6
      Top             =   6720
      Width           =   1815
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
      Left            =   5880
      TabIndex        =   5
      Top             =   6720
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6800
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UNIT SEARCH:"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Click the desire unit before editing!"
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
      TabIndex        =   19
      Top             =   1200
      Width           =   3990
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATION"
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      Left            =   6360
      TabIndex        =   15
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
      Left            =   7080
      TabIndex        =   14
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
      Left            =   6360
      TabIndex        =   13
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
      Left            =   7080
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   240
      Picture         =   "frmdestination.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
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
      TabIndex        =   11
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
Attribute VB_Name = "frmdestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addnum As Byte
Dim data_update As String
Private Sub cmdadd_Click()
addnum = 1
Me.Frame1.Visible = True
End Sub

Private Sub cmdcancel_Click()
With rsdest
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
End Sub

Private Sub cmddelete_Click()
On Error GoTo fixDel
With rsdest
        If .RecordCount > 0 Then
            If MsgBox("  Are You Sure You Want To Delete This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
             Call addNewLog(currentuser, "Delete Destination - Destination id: " + CStr(rsdest!ID) + " - Destination name: " + rsdest!Destination)
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
Me.txtdestination.Text = rsdest![Destination]
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdsave_Click()
data_update = dataUpdated()
Select Case addnum
 Case 1
 If rsdest.RecordCount > 0 Then
 rsdest.MoveFirst
While rsdest.EOF = False
If Trim$(Me.txtdestination.Text) = rsdest!Destination Then
MsgBox "Destination Name Already Exist!", vbOKOnly + vbCritical
Me.txtdestination.SetFocus
Exit Sub
End If
rsdest.MoveNext
Wend
End If
              With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tbldestination " & _
                "(destination)" & _
                " VALUES(" & _
                 "'" & Me.txtdestination.Text & "'" & _
                ")"
                  .Execute
                   Call addNewLog(currentuser, "Add Destination - Destination name: " + Me.txtdestination.Text)
            End With
            Call Emptyctl(Me, "txtcom")
       Call Form_Load
           
            MsgBox "   Record Successfully Add to the System.   ", vbOKOnly, "Success!"
Case 2
 With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tbldestination " & _
                "set destination='" & Me.txtdestination.Text & "'" & _
                " where id=" & rsdest!ID
              .Execute
               Call addNewLog(currentuser, "Update Destination - Destination id: " + CStr(rsdest!ID) + data_update)
            End With
            Call Emptyctl(Me, "txtcom")
                Call Form_Load
               
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
End Select
    Me.Frame1.Visible = False

End Sub



Private Sub Form_Load()
sDateString = Format(Now, "yyyy/mm/dd")
Me.Timer1.Enabled = True
Me.Timer1.Interval = 1000
Call dbconnect
End Sub

Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub

Private Sub dbconnect()
Set rsdest = Nothing
rsdest.Open "select * from tbldestination ", con, 3, 3
Set DataGrid1.DataSource = rsdest
    Call datasize
End Sub
Private Sub datasize()
With DataGrid1
            .Columns.Item(0).Visible = False
            .Columns.Item(1).width = 1500
End With
End Sub
Private Sub txtsearch_Change()
If Trim$(Me.txtsearch.Text) = "" Then
    Set rsdest = Nothing
    rsdest.Open "select * from tbldestination order by id ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rsdest
    Call datasize
      Set rsdest = Nothing
    rsdest.Open "select * from tbldestination order by id ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.txtsearch.Text) <> "" Then
Set rsdest = Nothing
    rsdest.Open "select * from tbldestination where destination like'%" & Me.txtsearch.Text & "%'", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rsdest
    Call datasize
End If
End Sub

Private Function dataUpdated() As String
If rsdest.RecordCount > 0 Then
If rsdest!Destination <> Me.txtdestination.Text Then
dataUpdated = dataUpdated + " - Destination name: " + rsdest!Destination
End If
End If
End Function




