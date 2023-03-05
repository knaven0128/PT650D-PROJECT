VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProduct 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commodity Registration"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "COMMODITY ENTRY"
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
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   13455
      Begin VB.TextBox txtprice 
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
         Text            =   "0.00"
         Top             =   1200
         Width           =   3615
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
         Left            =   11280
         TabIndex        =   22
         Top             =   4560
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
         Left            =   9360
         TabIndex        =   21
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txtproduct 
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
      Begin VB.TextBox txtdetails 
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
         Height          =   1695
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "txtcom"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE:"
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
         Left            =   2040
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMODITY:"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DETAILS:"
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
         TabIndex        =   3
         Top             =   2040
         Width           =   975
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
      TabIndex        =   20
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
      Left            =   5160
      TabIndex        =   19
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
      Left            =   3240
      TabIndex        =   18
      Top             =   6720
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
      TabIndex        =   17
      Top             =   6720
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
      Left            =   2880
      TabIndex        =   15
      Tag             =   "txtcom"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   2760
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COMMODITY SEARCH:"
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
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "COMMODITY REGISTRATION"
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
      TabIndex        =   14
      Top             =   0
      Width           =   11295
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   960
      Width           =   4530
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   480
      Picture         =   "frmProduct.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Click the desire commodity before editing!"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   4665
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
Attribute VB_Name = "frmProduct"
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
With rsproduct
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
With rsproduct
        If .RecordCount > 0 Then
            If MsgBox("  Are You Sure You Want To Delete This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
            Call addNewLog(currentuser, "Delete Commodity - commodity id: " + CStr(rsproduct!productid) + " - commodity: " + CStr(rsproduct!product_name))
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
Me.txtproduct.Text = rsproduct![product_name]
Me.txtprice.Text = rsproduct![product_price]
Me.txtdetails.Text = rsproduct![details]
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdsave_Click()
   update_data = dataUpdated()
If Me.txtproduct.Text = "" Then
    MsgBox "Don't Leave Product Empty!.", vbExclamation, "System Prompt"
     Me.txtproduct.SetFocus
    Exit Sub
       End If
If Me.txtdetails.Text = "" Then
    MsgBox "Don't Leave Details Empty.", vbExclamation, "System Prompt"
     Me.txtdetails.SetFocus
    Exit Sub
       End If
If Me.txtprice.Text = "0.00" Then
    MsgBox "Don't Leave Price Empty.", vbExclamation, "System Prompt"
     Me.txtprice.SetFocus
    Exit Sub
       End If
Select Case addnum
 Case 1
 If rsproduct.RecordCount > 0 Then
 rsproduct.MoveFirst
While rsproduct.EOF = False
If Trim$(Me.txtproduct.Text) = rsproduct!product_name Then
MsgBox "Product Name Already Exist!", vbOKOnly + vbCritical
Me.txtproduct.SetFocus
Exit Sub
End If
rsproduct.MoveNext
Wend
End If
              With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tblproduct " & _
                "(Product_Name,Product_Price,Details,Date_Added)" & _
                " VALUES(" & _
                 "'" & Me.txtproduct.Text & "'," & _
                 "'" & Me.txtprice.Text & "'," & _
                "'" & Me.txtdetails.Text & "'," & _
                "'" & sDateString & "'" & _
                ")"
                  .Execute
               Call addNewLog(currentuser, "Add Commodity - Commodity: " + Me.txtproduct.Text)
            End With
            Call Emptyctl(Me, "txtcom")
       Call Form_Load
           
            MsgBox "   Record Successfully Add to the System.   ", vbOKOnly, "Success!"
Case 2

 With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblproduct " & _
                "set Product_name='" & Me.txtproduct.Text & "'," & _
                "Product_Price='" & Me.txtprice.Text & "'," & _
                 "Details='" & Me.txtdetails.Text & "'" & _
                " where productid=" & rsproduct!productid
              .Execute
            
              Call addNewLog(currentuser, "Update Commodity - commodity id: " + CStr(rsproduct!productid) + update_data)
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
Call dbconnect
End Sub

Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub

Private Sub dbconnect()
Set rsproduct = Nothing
rsproduct.Open "select * from tblproduct ", con, 3, 3
Set DataGrid1.DataSource = rsproduct
    Call datasize
End Sub
Private Sub datasize()
With DataGrid1
            .Columns.Item(0).Visible = False
            .Columns.Item(1).width = 1500
            .Columns.Item(2).width = 1500
End With
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
Call numonly(KeyAscii)
End Sub
  
Private Sub txtsearch_Change()
If Trim$(Me.txtsearch.Text) = "" Then
    Set rsproduct = Nothing
    rsproduct.Open "select * from tblproduct order by productid ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rsproduct
    Call datasize
      Set rsproduct = Nothing
    rsproduct.Open "select * from tblproduct order by productid ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.txtsearch.Text) <> "" Then
Set rsproduct = Nothing
    rsproduct.Open "select * from tblproduct where Product_Name like'%" & Me.txtsearch.Text & "%'", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rsproduct
    Call datasize
End If
End Sub

Private Function dataUpdated() As String
If rsproduct.RecordCount > 0 Then
If rsproduct!product_name <> Me.txtproduct.Text Then
dataUpdated = dataUpdated + " - commodity: " + rsproduct!product_name
End If
If rsproduct!details <> Me.txtdetails.Text Then
dataUpdated = dataUpdated + " - details: " + rsproduct!details
End If

If rsproduct!product_price <> Me.txtprice.Text Then
dataUpdated = dataUpdated + " - price: " + CStr(rsproduct!product_price)
End If
End If
End Function
