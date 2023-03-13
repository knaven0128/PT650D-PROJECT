VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmComm 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmcomm.frx":0000
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   14490
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnegative 
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
      TabIndex        =   20
      Tag             =   "txtcom"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtpositive 
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
      TabIndex        =   18
      Tag             =   "txtcom"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000D&
      Caption         =   "Show Main Kilo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   65.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Left            =   4680
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   5280
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
      Left            =   12000
      TabIndex        =   15
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "&SET"
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
      Left            =   9960
      TabIndex        =   14
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox txtsymbol 
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
      TabIndex        =   12
      Tag             =   "txtcom"
      Top             =   3600
      Width           =   2415
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1200
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   5280
   End
   Begin VB.TextBox txtkilo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   65.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   9255
   End
   Begin VB.TextBox txtstr 
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
      TabIndex        =   5
      Tag             =   "txtcom"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtlen 
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
      TabIndex        =   4
      Tag             =   "txtcom"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtcomport 
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
      TabIndex        =   3
      Tag             =   "txtcom"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtcommsettings 
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
      TabIndex        =   2
      Tag             =   "txtcom"
      Text            =   "9600,N,8,1"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGATIVE SYMBOL:"
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
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "POSITIVE SYMBOL:"
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
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "UM SYMBOL:"
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
      Left            =   840
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   120
      Picture         =   "frmcomm.frx":0420
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1680
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   960
      Width           =   3690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL OF STRING:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL OF LENGHT:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPORT SETTINGS:"
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPORT NUMBER:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "CONNECTION SETTINGS"
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
      Left            =   1920
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
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim positive As String
Dim negative As String
Dim containPositive As Long
Dim containPositive2 As Long
Dim containNegative As Long
Dim containNegative2 As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
Me.Text1.Visible = True
Else
Me.Text1.Visible = False
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
Load MainForm
End Sub

Private Sub cmdset_Click()
On Error GoTo ShowError
 With cmd
                .ActiveConnection = ocn
                .CommandType = adCmdText
                .CommandText = "update tblcomm " & _
                "set Portnum='" & Me.txtcomport.Text & "'," & _
                 "commset='" & Me.txtcommsettings.Text & "'," & _
                 "comm_len='" & Me.txtlen.Text & "'," & _
                 "comm_str='" & Me.txtstr.Text & "'," & _
                  "comm_symbol='" & Me.txtsymbol.Text & "'," & _
                   "comm_positive='" & Me.txtpositive.Text & "'," & _
                    "comm_negative='" & Me.txtnegative.Text & "'" & _
                "where CommID=" & 1
              .Execute
            End With
            MsgBox "Data totally save!", vbInformation, "Info"
            Call Form_Load
Set rscomm = Nothing
 rscomm.Open "select * from tblcomm ", ocn, 3, 3
 comnum = rscomm![PortNum]
 comset = rscomm![Commset]
 commlen = rscomm![comm_Len]
 commstr = rscomm![comm_Str]
If (MSComm1.PortOpen = False) Then
MSComm1.settings = comset
MSComm1.InputLen = 0
Me.MSComm1.CommPort = comnum
MSComm1.RThreshold = 1
MSComm1.PortOpen = True
Timer1.Enabled = True
Timer1.Interval = 1
Timer2.Enabled = True
Timer2.Interval = 1
End If
Exit Sub
ShowError:
    Screen.MousePointer = vbDefault
    
    MsgBox "Error: " & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbExclamation
   
    Exit Sub
    
End Sub
Private Sub Form_Load()
On Error GoTo ShowError
Timer1.Enabled = False
Timer2.Enabled = False
Set rscomm = Nothing
rscomm.Open "select * from tblcomm", ocn, 3, 3
Me.txtcomport.Text = rscomm![PortNum]
Me.txtcommsettings.Text = rscomm![Commset]
Me.txtlen.Text = rscomm![comm_Len]
Me.txtstr.Text = rscomm![comm_Str]
Me.txtsymbol.Text = rscomm![comm_Symbol]
Me.txtpositive.Text = rscomm![comm_positive]
Me.txtnegative.Text = rscomm![comm_negative]
 comnum = rscomm![PortNum]
 comset = rscomm![Commset]
commsymbol = rscomm![comm_Symbol]
commpositive = rscomm![comm_positive]
commnegative = rscomm![comm_negative]
If MSComm1.PortOpen = False Then
MSComm1.settings = comset
MSComm1.InputLen = 0
Me.MSComm1.CommPort = comnum
MSComm1.RThreshold = 1
MSComm1.PortOpen = True
Timer1.Enabled = True
Timer1.Interval = 1
Timer2.Enabled = True
Timer2.Interval = 1
End If
ShowError:
Screen.MousePointer = vbDefault
MsgBox "Device is not Connected!", vbCritical, "Error"

    Exit Sub
End Sub



Private Sub Timer1_Timer()
Static WeightBuffer As String 'Create a permanent procedure level buffer
Dim Weight As String 'Temporary holding buffer
Dim prevWeight As String
Dim FinishPos As Long
Dim contain As Long
Select Case MSComm1.CommEvent 'Why was OnComm triggered?
Case comEvReceive 'OnComm was triggered because characters were received
    WeightBuffer = WeightBuffer & MSComm1.Input 'one or more characters were received, so concatenate them into buffer
    Do
  
        FinishPos = InStr(1, WeightBuffer, Me.txtsymbol.Text, vbTextCompare) 'is lb in our string?
        If FinishPos = 0 Then
           FinishPos = InStr(1, WeightBuffer, "lb", vbTextCompare) 'how about kg?
          FinishPos = InStr(1, WeightBuffer, "kg", vbTextCompare)
            FinishPos = InStr(1, WeightBuffer, "lg", vbTextCompare)
           FinishPos = InStr(1, WeightBuffer, "mg", vbTextCompare)
           FinishPos = InStr(1, WeightBuffer, "0€", vbTextCompare)
           FinishPos = InStr(1, WeightBuffer, " € ", vbTextCompare)
          FinishPos = InStr(1, WeightBuffer, "g", vbTextCompare)
        End If
        If FinishPos > 0 Then 'if we found either one then process it
            Weight = Left$(WeightBuffer, FinishPos + 1) 'put the piece we found in a temporary buffer
            WeightBuffer = Mid$(WeightBuffer, FinishPos + 2) 'store the unused data for future use
        Else
            Exit Do 'nothing found this loop so get out
        End If
    Loop
    If Len(Weight) > 0 Then 'Did we find anything to display?
        Me.Text1.Text = CStr(Mid(Weight, Val(Me.txtstr.Text), 40))
        
'        contain = InStr(1, Text1.Text, ")", vbTextCompare)
'        If contain = 0 Then
'        Me.Text1.Text = ""
'        End If
      
     ' Me.txtkilo.Text = CStr(Weight)
    End If
End Select
End Sub



Private Sub Timer2_Timer()
Dim current As String
positive = "+ "
negative = "- "

If (IsNumeric(Me.Text1.Text)) Then
current = Me.Text1.Text
End If

containPositive = InStr(1, Text1.Text, ")0", vbTextCompare)
containPositive2 = InStr(1, Text1.Text, ")8", vbTextCompare)

containNegative = InStr(1, Text1.Text, "):", vbTextCompare)
containNegative2 = InStr(1, Text1.Text, ")2", vbTextCompare)

If Trim(Me.Text1.Text) = "" Then
    Me.Text1.Text = Mid$(Me.Text1.Text, Val(txtstr.Text), Val(txtlen.Text))
    Else
    Me.txtkilo.Text = Mid$(Me.Text1.Text, Val(txtstr.Text), Val(txtlen.Text))
    Me.txtkilo.Text = ReturnNonAlpha(Me.txtkilo.Text)
    If containPositive > 0 Or containPositive2 > 0 Then
    Me.txtkilo.Text = positive + Me.txtkilo.Text
    ElseIf containNegative > 0 Or containNegative2 > 0 Then
    Me.txtkilo.Text = negative + Me.txtkilo.Text
    End If
    End If
End Sub

Public Function ReturnNonAlpha(ByVal sString As String) As String
   Dim i As Integer
   For i = 1 To Len(sString)
       If Mid(sString, i, 1) Like "[0-9]" Then
           ReturnNonAlpha = ReturnNonAlpha & Mid(sString, i, 1)
       End If
   Next i
End Function


