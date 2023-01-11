VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Registration"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13125
   ControlBox      =   0   'False
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13125
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "CUSTOMER ENTRY"
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
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   13335
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
         Left            =   11160
         TabIndex        =   18
         Top             =   5400
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
         Left            =   9240
         TabIndex        =   23
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtcontact 
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
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtaddres 
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
         TabIndex        =   2
         Tag             =   "txtcom"
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtfname 
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
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NUMBER:"
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
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME:"
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
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   1695
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
      Left            =   7080
      TabIndex        =   22
      Top             =   7800
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
      Left            =   5160
      TabIndex        =   21
      Top             =   7800
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
      Left            =   3240
      TabIndex        =   20
      Top             =   7800
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
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
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
      Left            =   2160
      TabIndex        =   16
      Tag             =   "txtcom"
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6588
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER SEARCH:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Click the desire customer before editing!"
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
      TabIndex        =   15
      Top             =   1200
      Width           =   4530
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
      TabIndex        =   14
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   360
      Picture         =   "frmCustomer.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
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
      TabIndex        =   7
      Top             =   1920
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
      TabIndex        =   6
      Top             =   1920
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
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
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
      TabIndex        =   4
      Top             =   1680
      Width           =   615
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER REGISTRATION"
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
      TabIndex        =   11
      Top             =   0
      Width           =   10215
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
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim update_data As String

Dim addnum As Byte
Dim printnum As Byte
Private Sub cmdadd_Click()
addnum = 1
Me.Frame1.Visible = True
End Sub

Private Sub cmdcancel_Click()
With rscustomer
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
With rscustomer
        If .RecordCount > 0 Then
            If MsgBox("  Are You Sure You Want To Delete This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
              Call addNewLog(currentuser, "Delete Customer - customer id: " + CStr(rscustomer!customerid) + "- customer name: " + rscustomer!customer_name)
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
Me.txtfname.Text = rscustomer![customer_name]
Me.txtaddres.Text = rscustomer![Address]
Me.txtcontact.Text = rscustomer![contact_number]
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdsave_Click()
update_data = dataUpdated()
If Me.txtfname.Text = "" Then
    MsgBox "Don't Leave Name Empty!.", vbExclamation, "System Prompt"
     Me.txtfname.SetFocus
    Exit Sub
       End If
If Me.txtaddres.Text = "" Then
    MsgBox "Don't Leave Address Empty.", vbExclamation, "System Prompt"
     Me.txtaddres.SetFocus
    Exit Sub
       End If
If Me.txtcontact.Text = "" Then
    MsgBox "Don't Leave Contact Empty.", vbExclamation, "System Prompt"
     Me.txtcontact.SetFocus
    Exit Sub
       End If
Select Case addnum
 Case 1
 If rscustomer.RecordCount > 0 Then
rscustomer.MoveFirst
While rscustomer.EOF = False
If Trim$(Me.txtfname.Text) = rscustomer!customer_name Then
MsgBox "Customer Name Already Exist!", vbOKOnly + vbCritical
Me.txtfname.SetFocus
Exit Sub
End If
rscustomer.MoveNext
Wend
End If
              With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tblcustomer " & _
                "(customer_name,address,contact_number,date_added)" & _
                " VALUES(" & _
                 "'" & Me.txtfname.Text & "'," & _
                "'" & Me.txtaddres.Text & "'," & _
                "'" & Me.txtcontact.Text & "'," & _
                "'" & sDateString & "'" & _
                ")"
                  .Execute
                  Call addNewLog(currentuser, "Add Customer - customer name: " + Me.txtfname.Text)
            End With
             
            Call Emptyctl(Me, "txtcom")
       Call Form_Load
           
            MsgBox "   Record Successfully Add to the System.   ", vbOKOnly, "Success!"
Case 2
 With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblcustomer " & _
                "set Customer_name='" & Me.txtfname.Text & "'," & _
                 "address='" & Me.txtaddres.Text & "'," & _
                 "contact_number='" & Me.txtcontact.Text & "'" & _
                " where customerid=" & rscustomer!customerid
              .Execute
               Call addNewLog(currentuser, "Update Customer - customer id: " + CStr(rscustomer!customerid) + update_data)
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
Set rscustomer = Nothing
rscustomer.Open "select * from tblcustomer ", con, 3, 3
Set DataGrid1.DataSource = rscustomer
    Call datasize
End Sub
Private Sub datasize()
With DataGrid1
If currentposition = "Weigher" Then
   .Columns.Item(0).Visible = False
End If
            .Columns.Item(1).width = 2500
            .Columns.Item(2).width = 2500
            .Columns.Item(3).width = 2000
End With
End Sub


Private Sub txtsearch_Change()
If Trim$(Me.txtsearch.Text) = "" Then
    Set rscustomer = Nothing
    rscustomer.Open "select * from tblcustomer order by customerid ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rscustomer
    Call datasize
      Set rscustomer = Nothing
    rscustomer.Open "select * from tblcustomer order by customerid ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.txtsearch.Text) <> "" Then
Set rscustomer = Nothing
    rscustomer.Open "select * from tblcustomer where customer_Name like'%" & Me.txtsearch.Text & "%'", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rscustomer
    Call datasize
End If
End Sub

Private Function dataUpdated() As String
If rscustomer.RecordCount > 0 Then
If rscustomer!customer_name <> Me.txtfname.Text Then
dataUpdated = dataUpdated + " - customer name: " + rscustomer!customer_name
End If
If rscustomer!Address <> Me.txtaddres.Text Then
dataUpdated = dataUpdated + " - address: " + rscustomer!Address
End If

If rscustomer!contact_number <> Me.txtcontact.Text Then
dataUpdated = dataUpdated + " - contact number: " + CStr(rscustomer!contact_number)
End If
End If
End Function
