VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDailyTransac 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Reports"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9045
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
      Left            =   6720
      TabIndex        =   20
      Top             =   3960
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
      Left            =   6720
      TabIndex        =   19
      Top             =   4680
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
      Left            =   6600
      TabIndex        =   18
      Top             =   2760
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
      Left            =   6600
      TabIndex        =   17
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   1800
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3960
      Width           =   6315
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   123928577
      CurrentDate     =   42922
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   1680
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   123928579
      CurrentDate     =   42922
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
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
      Left            =   2160
      TabIndex        =   15
      Top             =   840
      Width           =   5550
   End
   Begin VB.Label Label7 
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
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   3945
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
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
      Left            =   360
      TabIndex        =   9
      Top             =   4470
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE TO:"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE FROM:"
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
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
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
      Left            =   2760
      TabIndex        =   3
      Top             =   1440
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
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
      Left            =   5880
      TabIndex        =   1
      Top             =   1440
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   8415
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOM REPORT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   1680
      TabIndex        =   14
      Top             =   0
      Width           =   10215
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
      TabIndex        =   13
      Top             =   240
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   360
      Picture         =   "frmDailyTransac.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1815
      Index           =   1
      Left            =   -3960
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmDailyTransac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vdate As String
Dim mintCount As Integer, mintPause As Integer
Dim strDate As String
Dim fromdate As String
Dim todate As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdbrowse_Click()

Dim strTemp As String
    strTemp = fBrowseForFolder(Me.hWnd, "Select export path")
    If strTemp <> "" Then
    txtDestination = strTemp
    End If
End Sub

Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub

Private Sub cmdexport_Click()
On Error GoTo Err:
Set rsexport = Nothing
rsexport.Open "Select * from tblweighing where transaction_date>='" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and transaction_date<='" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid ", con, adOpenDynamic, adLockOptimistic
If Me.txtDestination.Text = "" Then
MsgBox "Select your file destination first!" & vbNewLine & "Please Click Browse.", vbOKOnly, "Warning"
ElseIf rsexport.RecordCount = 0 Then
MsgBox "No transaction in the Record has found!" & vbNewLine & "Please Contact the programmer.", vbOKOnly, "Warning"
Else
Me.txtDestination.Enabled = False
cmdexport.Enabled = False
Me.cmdbrowse.Enabled = False
Me.cmdclose.Enabled = False
Me.cmdprint.Enabled = False
Me.Timer2.Enabled = True
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
NumberOfRows = rsexport.RecordCount
rsexport.MoveFirst
For R = 1 To NumberOfRows
DataArray(R, 1) = rsexport.Fields("consec_no")
DataArray(R, 2) = rsexport.Fields("Plate_Number")
If rsexport!weigh_in > rsexport!weigh_out Then
DataArray(R, 3) = rsexport.Fields("Weigh_IN")
DataArray(R, 4) = rsexport.Fields("Weigh_Out")
Else
DataArray(R, 3) = rsexport.Fields("Weigh_Out")
DataArray(R, 4) = rsexport.Fields("Weigh_In")
End If
DataArray(R, 5) = rsexport.Fields("Net_Weight")
DataArray(R, 6) = rsexport.Fields("DateTime_WeighIN")
DataArray(R, 7) = rsexport.Fields("Datetime_weighOut")
DataArray(R, 8) = rsexport.Fields("Qty")
DataArray(R, 9) = rsexport.Fields("Unit")
DataArray(R, 10) = rsexport.Fields("Price")
DataArray(R, 11) = rsexport.Fields("TotalPrice")
DataArray(R, 12) = rsexport.Fields("Transaction_Date")
DataArray(R, 13) = rsexport.Fields("weigher")
DataArray(R, 14) = rsexport.Fields("customer_name")
DataArray(R, 15) = rsexport.Fields("product_name")
DataArray(R, 16) = rsexport.Fields("average")
DataArray(R, 17) = rsexport.Fields("scale_price")
DataArray(R, 18) = rsexport.Fields("Status")
DataArray(R, 19) = rsexport.Fields("Remarks")
DataArray(R, 20) = rsexport.Fields("destination")
rsexport.MoveNext
Next
Set oSheet = oBook.Worksheets(1)
oSheet.Range("A1:K1").Font.Bold = True
oSheet.Range("A1:A100").HorizontalAlignment = &HFFFFEFDD
oSheet.Range("A1:A100").NumberFormat = "@"
oSheet.Range("A2:K2").Font.Bold = True
oSheet.Range("A1:K1").Font.Size = 30
oSheet.Range("A2:K2").Font.Size = 16
oSheet.Range("C1").Value = rscompany!company_name
oSheet.Range("D2").Value = "(ALL TRANSACTION)"
oSheet.Range("A3").Value = "Date Printed:"
oSheet.Range("B3").Value = Me.Label9.Caption
oSheet.Range("A4").Value = "Time Printed:"
oSheet.Range("B4").Value = Me.Label11.Caption
oSheet.Range("A7:T7").Font.Bold = True
oSheet.Range("A7:T7").Borders.Weight = xlHairline
oSheet.Range("A2:M2").ColumnWidth = 20
oSheet.Range("A7:T7").Value = Array("Weighing Number", "Plate Number", "Gross", "Tare", "Net Weight", "Date/Time Weigh IN", "Date/Time Weigh Out", "Quantity", "Unit", "Price", "Total Price", "Transaction Date", "Operator Name", "Customer Name", "Commodity", "Average", "Scale Price", "Status", "Remarks", "Destination")
oSheet.Range("A8").Resize(NumberOfRows, 20).Value = DataArray
oBook.SaveAs Trim$(Me.txtDestination.Text) & "\TruckWeightDataReports-" & Format(Now, "mm-dd-yyyy") & ".xlsx"
oExcel.Quit
End If

Exit Sub
Err:
Timer2.Enabled = False
Me.cmdbrowse.Enabled = True
Me.cmdclose.Enabled = True
Me.cmdprint.Enabled = True
Me.cmdexport.Enabled = True
Me.txtDestination.Enabled = True
Call addNewLog(currentuser, "Export to Excel Custom Report")
End Sub

Private Sub cmdprint_Click()
     With Reports
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where transaction_date>='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and transaction_date<='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid"
                .Restart
                If rscompany![namecheck] = 0 Then
                .lblname.Visible = False
            Else
                .lblname.Visible = True
                .lblname.Caption = rscompany![company_name]
            End If
            If rscompany![emailcheck] = 0 Then
                .lblemail.Visible = False
            Else
                .lblemail.Visible = True
                .lblemail.Caption = rscompany![company_email]
            End If
            If rscompany![contactcheck] = 0 Then
                .lblcontact.Visible = False
            Else
                .lblcontact.Visible = True
                .lblcontact.Caption = rscompany![company_contact]
            End If
            If rscompany![addresscheck] = 0 Then
                .lbladdress.Visible = False
            Else
                .lbladdress.Visible = True
                .lbladdress.Caption = rscompany![company_address]
            End If
                .Label25.Caption = Format(Me.DTPicker1.Value, "yyyy/mm/dd") & " to " & Format(Me.DTPicker2.Value, "yyyy/mm/dd")
                .Label16.Caption = Me.Label9.Caption
                .Label17.Caption = Me.Label11.Caption
                .Show 1
                Call addNewLog(currentuser, "Print Custom Report")
                End With
End Sub




Private Sub Form_Load()
cmdexport.Enabled = True
strDate = Format(Now, "yyyy/mm/dd")
Me.DTPicker1.Value = strDate
Me.DTPicker2.Value = strDate
'Me.DTPicker1.Value = Format$(Me.DTPicker1.Value, "yyyy/mm/dd")
'Me.DTPicker2.Value = Format$(Me.DTPicker1.Value, "yyyy/mm/dd")
Set rscompany = Nothing
rscompany.Open "select * from tblcompany", ocn, 3, 3
Set rsexcel = Nothing
rsexcel.Open "select * from tblweighing", con, 3, 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load MainForm
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
   If MsgBox("Exporting data is Succesful!Open file destination?", vbYesNo + vbQuestion, "Successfully") = vbYes Then
        Cancel = 1
        ShellExecute 0, vbNullString, Me.txtDestination.Text, vbNullString, vbNullString, 1
    End If
    rsexport.Close
    progStat.Value = 0
    mintCount = 2
    Me.cmdbrowse.Enabled = True
    Me.cmdclose.Enabled = True
    Me.cmdprint.Enabled = True
    Me.txtDestination.Enabled = False
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
        mintCount = mintCount
        lblCount.Caption = "(" & mintCount & "%)..."
        progStat.Value = mintCount + 1
         
    ElseIf mintCount < 100 Then
        mintCount = mintCount + 2
        lblCount.Caption = "(" & mintCount & "%)..."
         progStat.Value = mintCount
        
    End If
    
    If mintPause = 100 Then
        lblCount.Caption = "(0%)..."
        lblInform.Caption = "Please wait the data is exporting......"
    ElseIf mintPause > 180 Then
   End If
End Sub

