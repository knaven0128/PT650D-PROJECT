VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDailyReports 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Reports"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   15675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&BROWSE...."
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
      Left            =   13800
      TabIndex        =   21
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdexport 
      Caption         =   "&EXPORT"
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
      Left            =   13800
      TabIndex        =   20
      Top             =   7800
      Width           =   1575
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
      Left            =   13080
      TabIndex        =   19
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&PRINT"
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
      Left            =   11040
      TabIndex        =   18
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   2040
   End
   Begin VB.TextBox txtDestination 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7080
      Width           =   6615
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmDailyReports.frx":0000
      Left            =   2040
      List            =   "frmDailyReports.frx":000D
      TabIndex        =   10
      Text            =   "Select here.........."
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Left            =   16080
      Top             =   2400
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6800
      _Version        =   393216
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
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click the export button to export data."
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
      Left            =   2640
      TabIndex        =   17
      Top             =   1320
      Width           =   3900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click the browse button to browse your file destination."
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
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   5550
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FILE DESTINATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblInform 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait the data is exporting......."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   7440
      TabIndex        =   14
      Top             =   7590
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "(0%)..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11025
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   12360
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT AND EXPORT:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:Select data you want to print/export!"
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
      TabIndex        =   7
      Top             =   840
      Width           =   4245
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   600
      Picture         =   "frmDailyReports.frx":004C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
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
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   1320
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   1320
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
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   960
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   6
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "EXPORT AND PRINT REPORTS"
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
      TabIndex        =   5
      Top             =   0
      Width           =   11655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Index           =   1
      Left            =   -240
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmDailyReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mintCount As Integer, mintPause As Integer
Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub

Private Sub cmdexport_Click()
If Me.txtDestination.Text = "" Then
MsgBox "Select you file destination first!", vbOKOnly, "Warning"
ElseIf Me.Combo1.Text = "Select here.........." Then
MsgBox "Select data you want to export!", vbOKOnly, "Information"
Me.Combo1.SetFocus
Else
Me.Timer2.Enabled = True
End If
If Me.Combo1.Text = "All Transaction" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing order by weighid ", con, adOpenDynamic, adLockOptimistic
Call exportall
ElseIf Me.Combo1.Text = "Daily Transaction" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN like '" & Me.Label1.Caption & "%' order by weighid ", con, adOpenDynamic, adLockOptimistic
Call dailyexport
Else
Me.Timer2.Enabled = True
End If
End Sub

Private Sub cmdprint_Click()
If Me.Combo1.Text = "All Transaction" Then
 With Reports
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing order by weighid"
                .Restart
                 .Label5.Caption = rscompany![company_name]
                .Label20.Caption = rscompany![company_address]
                .Label21.Caption = rscompany![company_contact]
                .Label20.Caption = "All Transaction Reports"
                .Label16.Caption = Me.Label9.Caption
                .Label17.Caption = Me.Label11.Caption
                .Show 1
            End With
ElseIf Me.Combo1.Text = "Daily Transaction" Then
            With Reports
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where DateTime_WeighIN like '" & Me.Label1.Caption & "%' order by weighid"
                .Restart
                 .Label5.Caption = rscompany![company_name]
                .Label20.Caption = rscompany![company_address]
                .Label21.Caption = rscompany![company_contact]
                .Label20.Caption = "Daily Transaction Reports"
                .Label16.Caption = Me.Label9.Caption
                .Label17.Caption = Me.Label11.Caption
                .Show 1
            End With
  End If
End Sub

Private Sub Combo1_Click()
If Me.Combo1.Text = "All Transaction" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing order by weighid ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
Call datasize
ElseIf Me.Combo1.Text = "Daily Transaction" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN like '" & Me.Label1.Caption & "%' order by weighid ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
Call datasize
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Set rscompany = Nothing
rscompany.Open "select * from tblcompany", con, 3, 3
sDateString = Format(Now, "m/d/yyyy")
Me.Label1.Caption = sDateString
Me.Timer1.Enabled = True
Me.Timer1.Interval = 1000
End Sub

Private Sub dailyexport()
On Error GoTo error
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Set rsexport = New ADODB.Recordset
rsexport.Open "Select * from tblweighing where DateTime_WeighIN like '" & Me.Label1.Caption & "%' order by weighid ", con, adOpenDynamic, adLockOptimistic
NumberOfRows = rsexport.RecordCount
rsexport.MoveFirst
For R = 1 To NumberOfRows
DataArray(R, 1) = rsexport.Fields("consec_no")
DataArray(R, 2) = rsexport.Fields("Plate_Number")
DataArray(R, 3) = rsexport.Fields("Weigh_IN")
DataArray(R, 4) = rsexport.Fields("Weigh_Out")
DataArray(R, 5) = rsexport.Fields("Net_Weight")
DataArray(R, 6) = rsexport.Fields("DateTime_WeighIN")
DataArray(R, 7) = rsexport.Fields("Datetime_weighOut")
DataArray(R, 8) = rsexport.Fields("weigher")
DataArray(R, 9) = rsexport.Fields("customer_name")
DataArray(R, 10) = rsexport.Fields("product_name")
rsexport.MoveNext
Next
Set oSheet = oBook.Worksheets(1)
oSheet.Range("A1:J1").Font.Bold = True
oSheet.Range("A2:J2").Font.Bold = True
oSheet.Range("A1:J1").Font.Size = 30
oSheet.Range("A2:J2").Font.Size = 16
oSheet.Range("C1").Value = rscompany![company_name]
oSheet.Range("E2").Value = "(DAILY REPORTS)"
oSheet.Range("H3").Value = "Date Printed:"
oSheet.Range("I3").Value = Me.Label9.Caption
oSheet.Range("H4").Value = "Time Printed:"
oSheet.Range("I4").Value = Me.Label11.Caption
oSheet.Range("A7:J7").Font.Bold = True
oSheet.Range("A2:J2").ColumnWidth = 12
oSheet.Range("A7:J7").Value = Array("Transaction Number", "Plate Number", "Weigh IN", "Weigh Out", "Net Weight", "Date/Time Weigh IN", "Date/Time Weigh Out", "weigher Name", "Customer Name", "Product Name")
oSheet.Range("A8").Resize(NumberOfRows, 10).Value = DataArray
oBook.SaveAs Trim(Me.txtDestination.Text) & "\DailyReports-" & Format(Now, "dd-mm-yyyy") & ".xls"
oExcel.Quit
rsexport.MoveFirst
Set rsexport = Nothing
error:
If Err.Number = 6160 Then
End If
End Sub
Private Sub exportall()
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Set rsexport = New ADODB.Recordset
rsexport.Open "Select * from tblweighing order by weighid ", con, adOpenDynamic, adLockOptimistic
NumberOfRows = rsexport.RecordCount
rsexport.MoveFirst
For R = 1 To NumberOfRows
DataArray(R, 1) = rsexport.Fields("consec_no")
DataArray(R, 2) = rsexport.Fields("Plate_Number")
DataArray(R, 3) = rsexport.Fields("Weigh_IN")
DataArray(R, 4) = rsexport.Fields("Weigh_Out")
DataArray(R, 5) = rsexport.Fields("Net_Weight")
DataArray(R, 6) = rsexport.Fields("DateTime_WeighIN")
DataArray(R, 7) = rsexport.Fields("Datetime_weighOut")
DataArray(R, 8) = rsexport.Fields("weigher")
DataArray(R, 9) = rsexport.Fields("customer_name")
DataArray(R, 10) = rsexport.Fields("product_name")
rsexport.MoveNext
Next
Set oSheet = oBook.Worksheets(1)
oSheet.Range("A1:J1").Font.Bold = True
oSheet.Range("A2:J2").Font.Bold = True
oSheet.Range("A1:J1").Font.Size = 30
oSheet.Range("A2:J2").Font.Size = 16
oSheet.Range("C1").Value = "VISAYAS COCO DEVELOPMENT, INC"
oSheet.Range("E2").Value = "(ALL TRANSACTION)"
oSheet.Range("H3").Value = "Date Printed:"
oSheet.Range("I3").Value = Me.Label9.Caption
oSheet.Range("H4").Value = "Time Printed:"
oSheet.Range("I4").Value = Me.Label11.Caption
oSheet.Range("A7:J7").Font.Bold = True
oSheet.Range("A2:J2").ColumnWidth = 12
oSheet.Range("A7:J7").Value = Array("Transaction Number", "Plate Number", "Weigh IN", "Weigh Out", "Net Weight", "Date/Time Weigh IN", "Date/Time Weigh Out", "weigher Name", "Customer Name", "Product Name")
oSheet.Range("A8").Resize(NumberOfRows, 10).Value = DataArray
oBook.SaveAs Trim(Me.txtDestination.Text) & "\ALLREPORTS-" & Format(Now, "dd-mm-yyyy") & ".xls"
oExcel.Quit
End Sub
Private Sub datasize()
With DataGrid1
            .Columns.Item(0).Visible = False
            .Columns.Item(1).Width = 1500
            .Columns.Item(2).Width = 1500
            .Columns.Item(3).Width = 1000
            .Columns.Item(4).Width = 1000
            .Columns.Item(5).Width = 1000
            .Columns.Item(6).Width = 1800
            .Columns.Item(7).Width = 2100
            .Columns.Item(8).Width = 2100
            .Columns.Item(9).Width = 1800
            .Columns.Item(10).Width = 1800
End With
End Sub


Private Sub mafl_Click()
frmlocation.Show 1
End Sub

Private Sub cmdbrowse_Click()
Dim strTemp As String
If Me.Combo1.Text = "Select here.........." Then
MsgBox "Select first the data you want to export!", vbOKOnly, "Warning"
ElseIf rstruck.RecordCount <= 0 Then
MsgBox "No data to export!", vbOKOnly, "Warning"
Else
    strTemp = fBrowseForFolder(Me.hWnd, "Select export path")
    If strTemp <> "" Then
    txtDestination = strTemp
    End If
End If
End Sub

Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub
Private Sub Timer2_Timer()
 Call CountMe
    lblCount.Visible = True
    lblInform.Visible = True
'    lblCBK.Visible = True
    progStat.Visible = True
    progStat.Value = progStat.Value + 2
   
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
If progStat.Value = 100 Then
    MsgBox "Exporting data is already complete.Thank You!", vbOKOnly, "Success!"
    Timer2.Enabled = False
    lblCount.Visible = False
    lblInform.Visible = False
   progStat.Visible = False
   Me.txtDestination.Text = ""
Call Form_Load
Else
    If txtDestination.Text = "" Then
     progStat.Value = 0
     
       'Your function, can be anything. Open another form, frmMain.show... Ect.
    End If
    End If
End Sub
Private Sub CountMe()
   mintPause = mintPause + 1
   
    If mintCount < 0 Then
        mintCount = mintCount + 1
        lblCount.Caption = "(" & mintCount & "%)..."
         
    ElseIf mintCount < 100 Then
        mintCount = mintCount + 2
        lblCount.Caption = "(" & mintCount & "%)..."
        
    End If
    
    If mintPause = 100 Then
        lblCount.Caption = "App..."
        lblInform.Caption = "Starting"
    ElseIf mintPause > 180 Then
   End If
End Sub


