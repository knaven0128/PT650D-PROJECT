VERSION 5.00
Begin VB.Form frmlocation 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOCATION"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin Weighing.jcbutton cmdsave 
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Save"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtlocate 
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
      Left            =   240
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   1800
      Width           =   5175
   End
   Begin Weighing.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   53
   End
   Begin Weighing.jcbutton jcbutton1 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Cancel"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   360
      Picture         =   "frmlocation.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
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
      TabIndex        =   4
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Paste in the field below the path of your excel export!"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   5850
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
Attribute VB_Name = "frmlocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc As String
Private Sub cmdsave_Click()
   With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tbllocation " & _
                "(Location,Date_added)" & _
                " VALUES(" & _
                "'" & loc & "'," & _
                "'" & sDateString & "'" & _
                ")"
                  .Execute
End With
            Call Emptyctl(Me, "txtcom")
            MsgBox "   Record Successfully Add to the System.   ", vbInformation, "Success!"
            Unload Me
            Load frmDailyReports
End Sub

Private Sub Form_Load()
Set rslocate = Nothing
rslocate.Open "Select * from tbllocation", con, 3, 3
sDateString = Format(Now, "mm/dd/yyyy")
loc = Trim(Me.txtlocate.Text)
End Sub

Private Sub jcbutton1_Click()
Unload Me
Load frmDailyReports
End Sub
