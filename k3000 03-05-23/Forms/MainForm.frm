VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10590
   ClientLeft      =   225
   ClientTop       =   -2445
   ClientWidth     =   17610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   10590
   ScaleWidth      =   17610
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   2400
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   1588
      ButtonWidth     =   2196
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TRUCK SCALE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CONNECTION"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CUSTOMER"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "COMMODITY"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DESTINATION"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "REPORTS"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "RESET DATA"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "USER'S"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "COMPANY"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LOG OUT"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BACK UP"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "POWER"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CONTACT US"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LOGS"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10215
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2478
            MinWidth        =   2478
            Picture         =   "MainForm.frx":29879
            Text            =   "Logged As:"
            TextSave        =   "Logged As:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2893
            MinWidth        =   2893
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2654
            MinWidth        =   2654
            Picture         =   "MainForm.frx":29E13
            Text            =   "Current User:"
            TextSave        =   "Current User:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Picture         =   "MainForm.frx":2A3AD
            Text            =   "Time:"
            TextSave        =   "Time:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1835
            MinWidth        =   1835
            TextSave        =   "6:34 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Picture         =   "MainForm.frx":2A947
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "3/5/2023"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2AF8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2E6B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":31BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":352E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3894F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3C04A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3F758
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":42DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":460E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":497FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4CDAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":501FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":53945
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":57085
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   11745
      Left            =   0
      Picture         =   "MainForm.frx":5A739
      Stretch         =   -1  'True
      Top             =   840
      Width           =   19200
   End
   Begin VB.Menu mt 
      Caption         =   "TRANSACTION"
      Begin VB.Menu mtf 
         Caption         =   "Trucks Information"
      End
      Begin VB.Menu mpi 
         Caption         =   "Commodity Information"
      End
      Begin VB.Menu mci 
         Caption         =   "Customer Information"
      End
   End
   Begin VB.Menu mu 
      Caption         =   "UTILITY"
      Begin VB.Menu mus 
         Caption         =   "User Maintenance"
      End
      Begin VB.Menu mscc 
         Caption         =   "Set Comm Connection"
      End
      Begin VB.Menu mdest 
         Caption         =   "Set Destination"
      End
      Begin VB.Menu msum 
         Caption         =   "Set Unit of Measure"
      End
      Begin VB.Menu mcp 
         Caption         =   "Set Company Profile"
      End
      Begin VB.Menu mdt 
         Caption         =   "Reset All Data"
      End
   End
   Begin VB.Menu mrs 
      Caption         =   "REPORTS"
      Begin VB.Menu mdtr 
         Caption         =   "Custom Transactions Report"
      End
   End
   Begin VB.Menu mp 
      Caption         =   "SYTEM PROGRAM"
      Begin VB.Menu mbud 
         Caption         =   "Back Up Database"
      End
      Begin VB.Menu mnp 
         Caption         =   "Note Pad"
      End
      Begin VB.Menu mcal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mlo 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mss 
         Caption         =   "Shutdown System"
      End
   End
   Begin VB.Menu ma 
      Caption         =   "ABOUT"
      Begin VB.Menu mpd 
         Caption         =   "Programmer Details"
      End
      Begin VB.Menu mlk 
         Caption         =   "License Key"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sdatestrings As String
Dim sTimeString As String
Dim copyright As String
Dim ddata As String


Private Sub Form_Load()
'Label7.Caption = "ALE INDUSTRIAL WEIGHING SCALE SERVICES, Technical Manager: ARNIEL L. ESPINAL, Address: BLk 144 L17 K80 Katapangan St. Karangalan Vill. San Isidro Cainta Rizal, Email: ale.scaleservice@gmail.com, Tel No.: (02) 508-62-71, Cell No.: (0977) 2999-931 (0916) 756-2337"

Set rscompany = Nothing
rscompany.Open "select * from tblcompany", ocn, 3, 3
StatusBar1.Panels(2).Text = currentposition
StatusBar1.Panels(4).Text = currentuser
StatusBar1.Panels(9).Text = rscompany!company_name
Call credits
Me.Caption = rscompany!company_name
'Label7.Caption = rscompanycompany_name + "                                                " + rscompany!company_name + "                                                " + rscompany!company_name
'Label7.Caption = Label7.Caption & Space(50)
End Sub

Private Sub Form_Resize()
   Image1.Top = 0
    Image1.Left = 0
    Image1.width = Me.ScaleWidth
    Image1.height = Me.ScaleHeight
'    Label7.Left = 0
'    Label7.width = Me.ScaleWidth
End Sub





Private Sub Label7_Click()

End Sub

Private Sub mbud_Click()
frmsqlBackup.Show 1
End Sub

Private Sub mcal_Click()
Shell "calc.exe"
End Sub



Private Sub mci_Click()
frmCustomer.Show 1
End Sub

Private Sub mcp_Click()
frmcompany.Show 1
End Sub

Private Sub mdest_Click()
frmdestination.Show 1
End Sub

Private Sub mdt_Click()
frmdeletedata.Show 1
End Sub

Private Sub mdtr_Click()
frmDailyTransac.Show 1
End Sub

Private Sub mlk_Click()
companyname = rscompany!company_name
copyright = Year(Now)
  Screen.MousePointer = vbDefault
    
   MsgBox "Registered to:" & vbNewLine & companyname & vbNewLine & rscompany!company_address & vbNewLine & "© " & copyright, vbOKOnly + vbInformation
End Sub

Private Sub mlo_Click()
If MsgBox("  Are You Sure You Log out?   ", vbQuestion + vbYesNo, "Notice") = vbYes Then
Unload Me
rsuser.Close
Load frmLogin
frmLogin.Show
     Call addNewLog(currentuser, "User LogOut")
Else
Cancel = 1
End If
End Sub

Private Sub mnp_Click()
Shell "notepad.exe"
End Sub

Private Sub mpd_Click()
frmprogrammer.Show 1
End Sub

Private Sub mpi_Click()
frmProduct.Show 1
End Sub

Private Sub mscc_Click()
frmComm.Show 1
End Sub

Private Sub mss_Click()
If MsgBox("  The system will now shutdown!", vbQuestion + vbYesNo, "Notice") = vbYes Then
End
Else
Cancel = 1
End If
End Sub

Private Sub msum_Click()
frmUM.Show 1
End Sub

Private Sub mtf_Click()
frmWT.Show 1
'frmWeighing.Show 1
End Sub

Private Sub mus_Click()
frmuseraccount.Show 1
End Sub

Private Sub Timer1_Timer()
'Dim str As String
'str = MainForm.Label7.Caption
'str = Mid$(str, 2, Len(str)) + Left(str, 1)
'MainForm.Label7.Caption = str
'Label7.Caption = VBA.Right(Label7.Caption, Len(Label7.Caption) - 1) + Left$(Label7.Caption, 1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.index
Case 1: frmWT.Show 1
'Case 1: frmWeighing.Show 1
Case 3: frmComm.Show 1
Case 4: frmCustomer.Show 1
Case 5: frmProduct.Show 1
Case 6: frmdestination.Show 1
Case 7: frmDailyTransac.Show 1
Case 8: frmdeletedata.Show 1
Case 9: frmuseraccount.Show 1
Case 10:
frmcompany.Show 1
Case 11: Call mlo_Click
Case 12: frmsqlBackup.Show 1
Case 13: Call mss_Click
Case 14: frmclient.Show 1
Case 15: frmLogs.Show 1
End Select
End Sub
Private Sub credits()
If currentposition = "Weigher" Then
Me.mus.Visible = False
Me.mcp.Visible = False
Me.mdt.Visible = False
mscc.Visible = False
Toolbar1.Buttons.Item(3).Visible = False
Toolbar1.Buttons.Item(7).Visible = False
Toolbar1.Buttons.Item(8).Visible = False
Toolbar1.Buttons.Item(9).Visible = False
Toolbar1.Buttons.Item(14).Visible = False
End If
If currentposition = "Administrator" Then
Me.mscc.Visible = False
Toolbar1.Buttons.Item(3).Visible = False
Toolbar1.Buttons.Item(9).Visible = False
End If
'If currentposition = "Seller" Then
'Me.mcp.Enabled = False
'Toolbar1.Buttons.Item(9).Enabled = False
'End If
End Sub
