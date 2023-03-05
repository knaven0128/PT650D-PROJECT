VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E9C7E0AE-FCCA-438C-B739-9E3133C371E8}#1.0#0"; "ToggleButtonActivex.ocx"
Begin VB.Form frmWeighing 
   BackColor       =   &H8000000D&
   Caption         =   "ALE INDUSTRIAL WEIGHING SCALE"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19260
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   19260
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   5160
   End
   Begin VB.Timer tdatetime 
      Interval        =   1000
      Left            =   0
      Top             =   6840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   6000
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   6000
   End
   Begin VB.Timer tcomm 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   120
   End
   Begin VB.CommandButton cmdin 
      Caption         =   "&IN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MouseIcon       =   "frmWeighing.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmWeighing.frx":0152
      TabIndex        =   78
      Top             =   960
      Width           =   1300
   End
   Begin VB.CommandButton cmdout 
      Caption         =   "&OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MouseIcon       =   "frmWeighing.frx":1099C
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   960
      Width           =   1300
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      MouseIcon       =   "frmWeighing.frx":10AEE
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   960
      Width           =   1305
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&PRINT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      MouseIcon       =   "frmWeighing.frx":10C40
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   960
      Width           =   1300
   End
   Begin VB.CommandButton cmdviewall 
      Caption         =   "VIEW ALL CO&MPLETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      MouseIcon       =   "frmWeighing.frx":10D92
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   960
      Width           =   1300
   End
   Begin VB.CommandButton cmdviewin 
      Caption         =   "VIEW &ALL INBOUND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      MouseIcon       =   "frmWeighing.frx":10EE4
      MousePointer    =   99  'Custom
      TabIndex        =   73
      Top             =   960
      Width           =   1300
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
      Height          =   855
      Left            =   9600
      MouseIcon       =   "frmWeighing.frx":11036
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   960
      Width           =   1300
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   0
      TabIndex        =   47
      Top             =   1920
      Width           =   19095
      Begin VB.TextBox txtkilo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   69.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1695
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0"
         Top             =   1080
         Width           =   9015
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   450
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Enter Plate Number:"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtplatenum 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   570
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txttransac 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblin 
         BackStyle       =   0  'Transparent
         Caption         =   "Weigh In:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   2040
         TabIndex        =   71
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblstatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Stable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   15960
         TabIndex        =   70
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbldate 
         BackStyle       =   0  'Transparent
         Caption         =   "lbldate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   16080
         TabIndex        =   69
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lbltime 
         BackStyle       =   0  'Transparent
         Caption         =   "lbltime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   16080
         TabIndex        =   68
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label lblwi 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0kg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3240
         TabIndex        =   67
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblwo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0kg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   7440
         TabIndex        =   66
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblout 
         BackStyle       =   0  'Transparent
         Caption         =   "Weigh Out:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   6000
         TabIndex        =   65
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblnw 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0kg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   11640
         TabIndex        =   64
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblnet 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   10200
         TabIndex        =   63
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lbldatewi 
         BackStyle       =   0  'Transparent
         Caption         =   "lbldate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   8520
         TabIndex        =   62
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lbldi 
         BackStyle       =   0  'Transparent
         Caption         =   " DATE/TIME OF IN:"
         DataField       =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   6480
         TabIndex        =   61
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbldo 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE/TIME OF OUT:"
         DataField       =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   6360
         TabIndex        =   60
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbldatewo 
         BackStyle       =   0  'Transparent
         Caption         =   "lbldate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   8520
         TabIndex        =   59
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "______________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H0000C000&
         Height          =   495
         Left            =   10680
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "GO!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   11640
         TabIndex        =   57
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1695
         Left            =   14160
         TabIndex        =   56
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblaverage 
         BackStyle       =   0  'Transparent
         Caption         =   "Average:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   14520
         TabIndex        =   55
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblavg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   15840
         TabIndex        =   54
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblsymbol 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   75.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1695
         Left            =   4200
         TabIndex        =   53
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.TextBox txtstatus 
      Height          =   285
      Left            =   3000
      TabIndex        =   46
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtdate 
      Height          =   285
      Left            =   3840
      TabIndex        =   45
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtmainkilo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   11160
      TabIndex        =   43
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Timer tmaincomm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   120
   End
   Begin VB.TextBox txtmidkilo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   15480
      TabIndex        =   42
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer tcommmid 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   240
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   240
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8160
      Top             =   240
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5415
      Left            =   840
      TabIndex        =   0
      Top             =   5640
      Width           =   16095
      Begin VB.CommandButton cmddest 
         Caption         =   "&ADD?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MouseIcon       =   "frmWeighing.frx":11188
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox txtweigher 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3960
         Width           =   3975
      End
      Begin VB.CommandButton cmdac 
         Caption         =   "&ADD?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         MouseIcon       =   "frmWeighing.frx":112DA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   720
         Width           =   705
      End
      Begin VB.CommandButton cmdap 
         Caption         =   "&ADD?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         MouseIcon       =   "frmWeighing.frx":1142C
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1200
         Width           =   705
      End
      Begin VB.CommandButton cmdau 
         Caption         =   "&ADD?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MouseIcon       =   "frmWeighing.frx":1157E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2640
         Width           =   1065
      End
      Begin VB.TextBox txtprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   10200
         TabIndex        =   9
         Text            =   "0"
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txttotalprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3000
         Width           =   3975
      End
      Begin VB.TextBox txtscaleprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   10200
         TabIndex        =   7
         Text            =   "0"
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox cmbcustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   2775
            Left            =   0
            TabIndex        =   5
            Top             =   45
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   0   'False
            HeadLines       =   1
            RowHeight       =   19
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2775
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4895
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   0   'False
            HeadLines       =   1
            RowHeight       =   19
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
      End
      Begin VB.CommandButton cmdselect 
         BackColor       =   &H000000FF&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaskColor       =   &H000000FF&
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin ToggleButtonActivex.ToggleButton ToggleButton1 
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   -2147483635
         BeginProperty OnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty OffFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtremarks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtweighid 
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
         Height          =   375
         Left            =   5520
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox cmbproduct 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2640
         TabIndex        =   19
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbdest 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         ItemData        =   "frmWeighing.frx":116D0
         Left            =   2640
         List            =   "frmWeighing.frx":116D2
         TabIndex        =   20
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   2640
         TabIndex        =   14
         Text            =   "0"
         Top             =   2160
         Width           =   4215
      End
      Begin VB.ComboBox cmbunit 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         ItemData        =   "frmWeighing.frx":116D4
         Left            =   2640
         List            =   "frmWeighing.frx":116D6
         TabIndex        =   15
         Top             =   2640
         Width           =   4215
      End
      Begin VB.Label lblcount 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   41
         Top             =   4080
         Width           =   5895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   40
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: After TURN ON Input the weight in the gray box."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   39
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SCALE OFFLINE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to Scale Offline."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Click ADD button to add new COMMODITY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   36
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " UNIT OF MEASURE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   35
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " QUANTITY OF:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   34
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMODITY :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   33
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATOR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   31
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   9120
         TabIndex        =   30
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PRICE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   3
         Left            =   8520
         TabIndex        =   29
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "T.S. PRICE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   5
         Left            =   8760
         TabIndex        =   28
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "WEIGHING NUMBER:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   27
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Click ADD button to add new CUSTOMER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Left            =   10200
         TabIndex        =   26
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Search, Arrow Down to select customer and Press Enter."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   25
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Click ADD button to add new UNIT OF MEASURE."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   24
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   23
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "COUNT # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   22
         Top             =   4080
         Width           =   2055
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   1320
      TabIndex        =   85
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2760
      TabIndex        =   84
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3840
      TabIndex        =   83
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F9)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5520
      TabIndex        =   82
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   6960
      TabIndex        =   81
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F6)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   8400
      TabIndex        =   80
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(ESC)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   9720
      TabIndex        =   79
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmWeighing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim savenum As Byte
Dim getnum As Byte
Dim printnum As Byte
Dim sdatetime As Date
Dim stime As Date
Dim rscount As ADODB.Recordset
Dim count_num As String
Private InitialControlList() As ControlInitial
Dim list_item As ListItem
Dim lst1 As ListItem
Dim X As New Class1
Dim i As Integer
Dim editnum As Integer
Dim customerKey As Integer
Dim productKey As Integer
Dim symbolize As String
Dim statusKilo As String
Dim current As String
Dim deproduct As String
Dim searchNow As Boolean




Private Sub cmbcustomer_KeyPress(KeyAscii As Integer)
  If KeyAscii = customerKey Then
       Call DataGrid1_Click
      
    End If
End Sub
Private Sub cmbproduct_KeyPress(KeyAscii As Integer)
  If KeyAscii = productKey Then
       Call DataGrid2_Click
    End If
End Sub

Private Sub cmbproduct_Change()
customerKey = 0
productKey = 13
If searchNow = True Then
If Trim$(Me.cmbproduct.Text) = "" And Trim$(Me.cmbproduct.Text) = "NA" Then
    Set rsproduct = Nothing
    rsproduct.Open "select * from tblproduct order by productid ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid2.DataSource = rsproduct
'      Set rsproduct = Nothing
'    rsproduct.Open "select * from tblproduct order by productid ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.cmbproduct.Text) <> "" Then
    Set rsproduct = Nothing
    rsproduct.Open "select product_name from tblproduct where product_name like'%" & Me.cmbproduct.Text & "%'", con, adOpenDynamic, adLockOptimistic
     searchNow = True
     
     With DataGrid2
    Set .DataSource = rsproduct
       .Columns.Item(0).width = 5000
    End With

    End If
    End If
End Sub

'Private Sub cmbcustomer_DropDown()
'cmbcustomer.Clear
'cmbcustomer.AddItem "None"
'Set rscustomer = Nothing
'With rscustomer
'.Open "Select * from tblcustomer", con, 3, 3
'Do Until .EOF
'cmbcustomer.AddItem !customer_name
'.MoveNext
'Loop
'End With
'rscustomer.Close
'End Sub

'Private Sub cmbproduct_DropDown()
'Me.cmbproduct.Clear
'cmbproduct.AddItem "NA"
'Set rsproduct = Nothing
'With rsproduct
'.Open "Select * from tblproduct", con, 3, 3
'Do Until .EOF
'cmbproduct.AddItem !product_name
'.MoveNext
'Loop
'End With
'rsproduct.Close
'End Sub

Private Sub cmbunit_DropDown()
Me.cmbunit.Clear
cmbunit.AddItem "NA"
Set rsum = Nothing
With rsum
.Open "Select * from tblunitmeasure", con, 3, 3
Do Until .EOF
cmbunit.AddItem !unit_name
.MoveNext
Loop
End With
rsum.Close
End Sub

Private Sub cmbdest_DropDown()
Me.cmbdest.Clear
cmbdest.AddItem "NA"
Set rsdest = Nothing
With rsdest
.Open "Select * from tbldestination", con, 3, 3
Do Until .EOF
cmbdest.AddItem !Destination
.MoveNext
Loop
End With
rsdest.Close
End Sub

Private Sub cmdac_Click()
frmCustomer.Show 1
End Sub

Private Sub cmdap_Click()
frmProduct.Show 1
End Sub

Private Sub cmdau_Click()
frmUM.Show 1
End Sub
Private Sub cmdclose_Click()
If MsgBox("  Are You Sure You Want To Close Transaction?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
Unload Me
Else
Cancel = 1
End If
End Sub


Private Sub cmdget_Click()

End Sub


Private Sub cmddest_Click()
frmdestination.Show 1
End Sub

Private Sub cmdin_Click()
If Me.lblstatus.Caption = "Unstable" Then
MsgBox "Can't capture Negative Value OR Weight Unstable!" & vbNewLine & "Please Contact the Programmer", vbOKOnly + vbInformation
 Exit Sub
End If
If lblsymbol.Caption = "-" Then
MsgBox "Can't capture negative value!" & vbNewLine & "Please Contact the Programmer", vbOKOnly + vbInformation

Else
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where plate_number = '" & Me.txtplatenum.Text & "' And status='IN'", con, 3, 3
Set rscompany = Nothing
rscompany.Open "Select * from tblcompany", ocn, 3, 3



If Me.txtkilo.Text = "" Then
MsgBox "The Scale is Empty!" & vbNewLine & "You can Scale Offline" & vbNewLine & "Contact the Programmer", vbOKOnly + vbInformation
ElseIf ReturnNonAlpha(Me.txtkilo.Text) < 1 Then
MsgBox "The Scale is 0!" & vbNewLine & "You cannot weight in 0 value" & vbNewLine & "Contact the Programmer", vbOKOnly + vbInformation
ElseIf Me.txtplatenum.Text = "" Then
MsgBox "Please Input the Plate Number", vbInformation, "Warning"
Me.txtplatenum.SetFocus
ElseIf rstruck.RecordCount = 1 Then
MsgBox "Plate Number Already Exist!"
Else
 Set rs = Nothing
    rs.Open "tblsetup", con, adOpenStatic, adLockReadOnly
        consec_num = generate((Int(rs!soldidcnt) + 1), 5)
            Me.txttransac.Text = consec_num
Call enabledall
Me.Text1.Text = 1
Me.Timer1.Enabled = True
Me.Timer2.Enabled = True
Me.Timer3.Enabled = False
Me.Shape1.Visible = True
Me.txtplatenum.SetFocus
Me.txttotalprice.Enabled = True
 Me.txtweigher.Text = currentname
If Me.cmbproduct.Text = "" Or Me.cmbcustomer.Text = "" Then
Me.cmbproduct.Text = rscompany!de_product
Me.cmbunit.Text = rscompany!de_unit
Me.cmbdest.Text = "NA"
End If
savenum = 1
printnum = 1
Me.cmdsave.Enabled = True
        Set rs = Nothing
    rs.Open "tblcount", con, adOpenStatic, adLockReadOnly
        count_num = generate((Int(rs!countnumber) + 1), 7)
        Me.lblcount.Caption = count_num
Me.lbldi.Visible = True
Me.lbldatewi.Visible = True
lblwi.Caption = ReturnNonAlpha(Me.txtkilo.Text) + " kg"
Me.lbldatewi.Caption = Me.lbldate + " " + Me.lbltime
Timer1.Enabled = False
Timer2.Enabled = False
Me.lblin.Visible = True
End If
End If

End Sub

Private Sub cmdout_Click()
If Me.lblstatus.Caption = "Unstable" Then
MsgBox "Can't capture Negative Value OR Weight Unstable!" & vbNewLine & "Please Contact the Programmer", vbOKOnly + vbInformation
 Exit Sub
End If

If ReturnNonAlpha(Me.txtkilo.Text) < 1 Then
MsgBox "The Scale is 0!" & vbNewLine & "You cannot weight in 0 value" & vbNewLine & "Contact the Programmer", vbOKOnly + vbInformation
 Exit Sub
End If
If lblsymbol.Caption = "-" Then
MsgBox "Can't capture negative value!" & vbNewLine & "Please Contact the Programmer", vbOKOnly + vbInformation
Else
Set rs = Nothing
rs.Open "Select * from tblweighing", con, 3, 3
Set rscompany = Nothing
rscompany.Open "Select * from tblcompany", ocn, 3, 3
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where plate_number = '" & Me.txtplatenum.Text & "' and status='IN'", con, 3, 3
If Me.txtkilo.Text = "" Then
MsgBox "The Scale is Empty!" & vbNewLine & "You can Scale Offline" & vbNewLine & "Contact the Programmer", vbOKOnly + vbInformation
ElseIf Me.txtplatenum.Text = "" Then
MsgBox "Please Input the Plate Number", vbInformation, "Warning"
Me.txtplatenum.SetFocus
ElseIf rstruck.RecordCount = 0 Then
MsgBox "Can't find " & Trim$(Me.txtplatenum.Text) & " in the record!" & vbNewLine & "Please Check Plate the Number! ", vbCritical, "Error"
Me.txtplatenum.SetFocus
Else
'cmdselect.Enabled = False
'Me.cmbcustomer.Enabled = False
'Me.cmbproduct.Enabled = False
If savenum = 1 Then
savenum = 2
printnum = 2
Me.cmbunit.Enabled = True
Me.cmbdest.Enabled = True
Me.txtqty.Enabled = True
Me.txtprice.Enabled = True
'Me.txttotalprice.Enabled = True
Me.Timer2.Enabled = False
Me.Timer3.Enabled = True
Me.Timer4.Enabled = True
Me.cmdac.Enabled = False
Me.cmdap.Enabled = False
Me.cmdau.Enabled = True
Me.txtscaleprice.Enabled = True
Me.txtplatenum.SetFocus
 Me.txtweigher.Text = currentname
 Me.cmbcustomer.Text = "Search Customer........."
Me.cmbproduct.Text = ""
Me.txtremarks.Enabled = True
    Me.txtweighid.Text = rstruck!weighid
    Me.txttransac.Text = rstruck!consec_no
    Me.txtweigher.Text = rstruck!weigher
    lblwi.Caption = rstruck!weigh_in & " kg"
    Me.lbldatewi.Caption = rstruck!datetime_weighin
    Me.cmbcustomer.Text = rstruck!customer_name
    Me.cmbproduct.Text = rstruck!product_name
    If Val(txtqty.Text) = 0 Then
    Me.txtqty.Text = rstruck!qty
    End If
    If Val(txtprice.Text) = 0 Then
    Me.txtprice.Text = rstruck!Price
    End If
    Me.cmbunit.Text = rstruck!unit
    Me.cmbdest.Text = rstruck!Destination & vbNullString
    Me.txtscaleprice.Text = rstruck!scale_price
    Me.txtremarks.Text = rstruck!Remarks
    Me.lblcount.Caption = rstruck!countnum
txtstatus.Text = "OUT"

End If
Timer3.Enabled = False
Timer4.Enabled = False
Me.lblout.Visible = True
lblwo.Caption = ReturnNonAlpha(Me.txtkilo.Text) & " kg"
lblnw.Caption = Val(lblwi.Caption) - Val(lblwo.Caption)
Me.lblnw.Caption = Me.lblnw.Caption & " kg"
Me.cmdsave.Enabled = True
Me.cmdprint.Enabled = False
If Val(Me.txtqty.Text) > 0 Then
lblavg.Caption = Val(lblnw.Caption) / Val(Me.txtqty.Text)
Me.lblavg.Caption = Format(Me.lblavg.Caption, "###,###,####,#.00")
Else
Me.lblavg.Caption = "0"
End If
Me.txttotalprice.Text = Val(Me.lblnw.Caption) * Val(Me.txtprice.Text)
End If
End If
End Sub

Private Sub cmdprint_Click()
           'If Check1.Value = 1 Then
           'RTruckScaleIN.Label5.Visible = False
           'Else
            ' RTruckScaleIN.Label5.Visible = True
           'End If
Select Case printnum
Case 1
    With OliverScaleIN
         .ado1.Connection = con
               .ado1.Source = "select * from tblweighing where consec_no ='" & Me.txttransac.Text & "'"
              .Restart
               .Label17.Caption = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
            If rscompany!addresscheck = 0 Then
                .lbladdress.Visible = False
            ElseIf rscompany!namecheck = 0 Then
                .lblcompanyname.Visible = False
            ElseIf rscompany!contactcheck = 0 Then
                .lblcontact.Visible = False
            ElseIf rscompany!emailcheck = 0 Then
                .lblemail.Visible = False
            Else
                .lbladdress.Visible = True
                .lbladdress.Caption = rscompany![company_address]
                .lblcompanyname.Visible = True
                .lblcompanyname.Caption = rscompany![company_name]
'                .lblcontact.Visible = True
                .lblcontact.Caption = rscompany![company_contact]
'                .lblemail.Visible = True
                .lblemail.Caption = rscompany![company_email]
            End If
               ' .Label14.Caption = Format(Now, "yyyy/mm/dd")
               ' .Label15.Caption = Format(Now, "hh:mm:ss AM/PM")
                .Show 1
                cmdsave.Enabled = False
                '.PrintReport False
Me.txtplatenum.SetFocus
Me.cmdin.Enabled = True
Me.cmdprint.Enabled = False
Me.txtqty.Text = "0"
Me.txtprice.Text = "0.00"
Me.cmbproduct.Text = ""
Me.cmbcustomer.Text = ""
Me.cmbproduct.Text = ""
Me.cmbunit.Text = "Select Unit........."
Me.cmbdest.Text = "Select Destination........."
Me.txtplatenum.Text = ""
Me.txtremarks.Text = ""
Me.lblwi.Caption = "0kg"
Me.lblavg.Caption = "0"
Me.txtweighid.Text = ""
Me.txtprice.Text = "0"
Me.txtscaleprice.Text = "0"
Me.txttotalprice.Text = "0.00"
Me.lbldi.Visible = False
Me.lbldatewi.Visible = False
Call disabled

         
            End With
           
Case 2
With OliverScaleOUT
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where consec_no ='" & Me.txttransac.Text & "'"
                .Restart
                  .Label17.Caption = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
            If rscompany!addresscheck = 0 Then
                .lbladdress.Visible = False
            ElseIf rscompany!namecheck = 0 Then
                .lblcompanyname.Visible = False
            ElseIf rscompany!contactcheck = 0 Then
                .lblcontact.Visible = False
            ElseIf rscompany!emailcheck = 0 Then
                .lblemail.Visible = False
            Else
'                .lbladdress.Visible = True
                .lbladdress.Caption = rscompany![company_address]
'                .lblcompanyname.Visible = True
                .lblcompanyname.Caption = rscompany![company_name]
'                .lblcontact.Visible = True
                .lblcontact.Caption = rscompany![company_contact]
'                .lblemail.Visible = True
                .lblemail.Caption = rscompany![company_email]
            End If
                .Show 1
                   cmdsave.Enabled = False
                '.PrintReport False
Me.txtplatenum.SetFocus
Me.cmdin.Enabled = True
Me.cmdprint.Enabled = False
Me.txtqty.Text = "0"
Me.cmbproduct.Text = ""
Me.cmbcustomer.Text = ""
Me.cmbproduct.Text = ""
Me.cmbunit.Text = "Select Unit........."
Me.cmbdest.Text = "Select Destination........."
Me.txtplatenum.Text = ""
Me.lblwi.Caption = "0kg"
Me.lblwo.Caption = "0kg"
Me.lblnw.Caption = "0kg"
Me.txtweighid.Text = ""
Me.txtremarks.Text = ""
Me.txtprice.Text = "0"
Me.txtscaleprice.Text = "0"
Me.lblavg.Caption = "0"
Me.txttotalprice.Text = "0.00"
Me.lbldi.Visible = False
Me.lbldatewi.Visible = False
Call disabled
End With
            End Select
End Sub

Private Sub cmdsave_Click()
Select Case savenum
 Case 1
 If Me.txtplatenum.Text = "" Then
    MsgBox "Please Input Plate Number!.", vbExclamation, "System Prompt"
     Me.txtplatenum.SetFocus
ElseIf Me.txtweigher.Text = "" Then
    MsgBox "Please Input weigher Name", vbExclamation, "System Prompt"
     Me.txtweigher.SetFocus
ElseIf Me.cmbcustomer.Text = "Search Customer........." Then
    MsgBox "Please Select Desire Customer Name.", vbExclamation, "System Prompt"
     Me.cmbcustomer.SetFocus
ElseIf Me.cmbcustomer.Text = "" Then
    MsgBox "Please Select Desire Customer Name.", vbExclamation, "System Prompt"
     Me.cmbcustomer.SetFocus
ElseIf Me.cmbproduct.Text = "" Then
    MsgBox "Please Select Desire Product Name.", vbExclamation, "System Prompt"
     Me.cmbproduct.SetFocus
ElseIf Me.cmbproduct.Text = "" Then
    MsgBox "Please Select Desire Product Name.", vbExclamation, "System Prompt"
     Me.cmbproduct.SetFocus
'ElseIf Me.cmbunit.Text = "Select Unit........." Then
'    MsgBox "Please Select Desire Unit.", vbExclamation, "System Prompt"
'     Me.cmbunit.SetFocus
'ElseIf Me.cmbdest.Text = "Select Destination........." Then
'    MsgBox "Please Select Desire Destination.", vbExclamation, "System Prompt"
'     Me.cmbdest.SetFocus
'     ElseIf Me.cmbdest.Text = "" Then
'    MsgBox "Please Select Desire Destination.", vbExclamation, "System Prompt"
'     Me.cmbdest.SetFocus
'ElseIf Me.cmbunit.Text = "" Then
'    MsgBox "Please Select Desire Unit.", vbExclamation, "System Prompt"
'     Me.cmbunit.SetFocus
     Else
        Me.cmdprint.Enabled = True
printnum = 1

txtstatus.Text = "IN"
              With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "INSERT INTO tblweighing " & _
                "(Consec_No,Plate_Number,weigher,transaction_date,Weigh_IN,Weigh_OUT,NET_WEIGHT,QTY,UNIT,Destination,Price,TotalPrice,DateTime_WeighIN,Customer_name,Product_name,Average,Scale_Price,Status,Remarks,Countnum )" & _
                " VALUES(" & _
                "'" & Me.txttransac.Text & "'," & _
                "'" & Me.txtplatenum.Text & "'," & _
                "'" & currentname & "'," & _
                "'" & Format(Me.lbldate.Caption, "yyyy/MM/dd") & "'," & _
                "'" & Val(lblwi.Caption) & "'," & _
                "'" & Val(lblwo.Caption) & "'," & _
                "'" & Val(lblnw.Caption) & "'," & _
                "'" & Val(Me.txtqty.Text) & "'," & _
                "'" & Me.cmbunit.Text & "'," & _
                "'" & Me.cmbdest.Text & "'," & _
                "'" & CDbl(Me.txtprice.Text) & "'," & _
                "'" & CDbl(Me.txttotalprice.Text) & "'," & _
                "'" & Format(Me.lbldatewi.Caption, "yyyy/mm/dd HH:MM:SS") & "'," & _
                "'" & Me.cmbcustomer.Text & "'," & _
                "'" & Me.cmbproduct.Text & "'," & _
                "'" & Val(lblavg.Caption) & "'," & _
                "'" & CDbl(Me.txtscaleprice.Text) & "'," & _
                "'" & txtstatus.Text & "'," & _
                "'" & Me.txtremarks.Text & "'," & _
                "'" & Me.lblcount.Caption & "'" & _
                ")"
                  .Execute
                          Call addNewLog(currentuser, "Weigh IN - Transaction No.: " + Me.txttransac.Text + " - Plate Number: " + Me.txtplatenum.Text)
            End With
            con.BeginTrans
        Call cmmd("update tblsetup set soldidcnt='" & consec_num & "'")
        Call cmmd("update tblcount set countnumber='" & count_num & "'")

            con.CommitTrans
            Call cmdprint_Click
            Call Emptyctl(Me, "txtcom")
            
      End If
Case 2
If Me.txtqty.Text = "" Then
 MsgBox "Qty Field can't be empty!,0 value accepted.", vbExclamation, "System Prompt"
 Me.txtqty.SetFocus
 ElseIf Me.txtprice.Text = "" Then
MsgBox "Price Field can't be empty!.0 value accepted", vbExclamation, "System Prompt"
Me.txtprice.SetFocus
 ElseIf Me.txtscaleprice.Text = "" Then
MsgBox "TS Price Field can't be empty!.0 value accepted", vbExclamation, "System Prompt"
Me.txtscaleprice.SetFocus
Else
Me.txttotalprice.Text = Val(Me.lblnw.Caption) * Val(Me.txtprice.Text)
If Val(Me.txtqty.Text) > 0 Then
lblavg.Caption = Val(lblnw.Caption) / Val(Me.txtqty.Text)
Me.lblavg.Caption = Format(Me.lblavg.Caption, "###,###,####,#.00")
Else
Me.lblavg.Caption = "0"
End If
   Me.cmdprint.Enabled = True
With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                "set Consec_No='" & Me.txttransac.Text & "'," & _
                 "Weigh_OUT='" & Val(lblwo.Caption) & "'," & _
                "DateTime_WeighOUT='" & Format$(Me.lbldate.Caption, "yyyy/mm/dd") & " " & Format$(Me.lbltime.Caption, "hh:mm:ss") & "'," & _
                "Net_Weight='" & Val(lblnw.Caption) & "'," & _
                "QTY='" & Val(Me.txtqty.Text) & "'," & _
                "UNIT='" & Me.cmbunit.Text & "'," & _
                "Destination='" & Me.cmbdest.Text & "'," & _
                "PRICE='" & CDbl(txtprice.Text) & "'," & _
                "Totalprice='" & CDbl(Me.txttotalprice.Text) & "'," & _
                "customer_name='" & Me.cmbcustomer.Text & "'," & _
                "Product_name='" & Me.cmbproduct.Text & "'," & _
                "Average='" & lblavg.Caption & "'," & _
                "Scale_Price='" & CDbl(Me.txtscaleprice.Text) & "'," & _
                "Status='" & Me.txtstatus.Text & "'," & _
                "Remarks='" & Me.txtremarks.Text & "'" & _
                " where weighID=" & rstruck!weighid
              .Execute
                         Call addNewLog(currentuser, "Weigh OUT - Transaction No.: " + Me.txttransac.Text + "- Plate Number: " + Me.txtplatenum.Text)
            End With
           Call cmdprint_Click

            
End If
      End Select
    If MSComm1.PortOpen = False Then
    If Me.ToggleButton1.Value = True Then
    Me.ToggleButton1.Value = False
    Me.txtkilo.Text = 0
    End If
    End If
End Sub


Private Sub cmdselect_Click()
Me.Frame4.height = 3000
      searchNow = True
Me.Frame4.Visible = True

  Set rsproduct = Nothing
    rsproduct.Open "select * from tblproduct order by productid ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid2.DataSource = rsproduct
      DataGrid2.Columns.Item(0).width = 0
        DataGrid2.Columns.Item(1).width = 5000
      DataGrid2.Columns.Item(2).width = 0
      DataGrid2.Columns.Item(3).width = 0
       DataGrid2.Columns.Item(4).width = 0
' If savenum = 2 Or savenum = 1 And Me.cmbproduct.Text <> "" Then
'    Me.Frame4.Visible = False
'    Else
'      Me.Frame4.Visible = True
'    End If
'End If
End Sub

Private Sub cmdviewall_Click()
listnum = 2
frmlist.Show 1
End Sub

Private Sub cmdviewin_Click()
listnum = 1
frmlist.Show 1

End Sub

Private Sub DataGrid1_Click()
 customerKey = 0
If rscustomer.RecordCount > 0 Then
cmbcustomer.Text = rscustomer!customer_name
Me.Frame2.Visible = False
Me.cmbcustomer.SetFocus
Else
Me.Frame2.Visible = False
End If
End Sub


Private Sub DataGrid2_Click()
productKey = 0
If rsproduct.RecordCount > 0 Then
cmbproduct.Text = rsproduct!product_name
Me.Frame4.Visible = False
Me.cmbproduct.SetFocus
Else
Me.Frame4.Visible = False
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = customerKey Then
Call DataGrid1_Click
End If
End Sub
Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = productKey Then
Call DataGrid2_Click
End If
End Sub

Private Sub Form_Load()
If currentposition = "Weigher" Or currentposition = "Administrator" Then
Me.txtmainkilo.Visible = False
Else
Me.txtmainkilo.Visible = True
End If
savenum = 1
printnum = 1
InitialControlList = GetLocation(Me)
ReSizePosForm Me, Me.height, Me.width, Me.Left, Me.Top, True
On Error GoTo ShowError
Set rs = Nothing
rs.Open "tblcount", con, 3, 3
Me.lblcount.Caption = rs!countnumber
Set rscomm = Nothing
rscomm.Open "select * from tblcomm ", ocn, 3, 3
comnum = rscomm!PortNum
comset = rscomm!Commset
commlen = rscomm!comm_Len
commstr = rscomm!comm_Str
commsymbol = rscomm!comm_Symbol
commpositive = rscomm!comm_positive
commnegative = rscomm!comm_negative
Me.txtweigher.Text = currentname
Me.cmbproduct.Text = ""
Me.cmbunit.Text = "Select Unit........."
Me.cmbdest.Text = "Select Destination........."
If MSComm1.PortOpen = False Then
MSComm1.settings = comset
MSComm1.InputLen = 0
Me.MSComm1.CommPort = comnum
MSComm1.RThreshold = 1
MSComm1.PortOpen = True
Me.tcomm.Enabled = True
Me.tcomm.Interval = 1
Me.tcommmid.Enabled = True
Me.tcommmid.Interval = 1
tmaincomm.Enabled = True
Me.tmaincomm.Interval = 1

End If
Exit Sub
ShowError:
   Screen.MousePointer = vbDefault
    
MsgBox "Device is not Connected! " & vbnextline & "Please scale offline", vbCritical, "Error"
   
    Exit Sub
If Val(Me.txtkilo.Text) <= 100 Then
Me.Shape1.BackColor = &HC000&
Me.Label21.Caption = "GO!"
Else
Me.Shape1.BackColor = vbRed
Me.Label21.Caption = "STOP!"
End If

End Sub

Private Sub Form_Resize()
ResizeControls Me, InitialControlList, True
End Sub

Private Sub lblnw_Change()
lblnw.Caption = Str(Abs(Val(lblnw.Caption))) + " kg"
End Sub


Private Sub tcomm_Timer()
Static WeightBuffer As String 'Create a permanent procedure level buffer
Dim Weight As String 'Temporary holding buffer
Dim FinishPos As Long
Dim contain As Long
Select Case MSComm1.CommEvent 'Why was OnComm triggered?
Case comEvReceive 'OnComm was triggered because characters were received
    WeightBuffer = WeightBuffer & MSComm1.Input 'one or more characters were received, so concatenate them into buffer
    Do
  
        FinishPos = InStr(1, WeightBuffer, commsymbol, vbTextCompare) 'is lb in our string?
        If FinishPos = 0 Then
           FinishPos = InStr(1, WeightBuffer, "lb", vbTextCompare) 'how about kg?
          FinishPos = InStr(1, WeightBuffer, "kg", vbTextCompare)
            FinishPos = InStr(1, WeightBuffer, "lg", vbTextCompare)
           FinishPos = InStr(1, WeightBuffer, "mg", vbTextCompare)
           FinishPos = InStr(1, WeightBuffer, "€", vbTextCompare)
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
        Me.txtmainkilo.Text = CStr(Mid(Weight, Int(commstr), Int(commlen)))
        contain = InStr(1, txtmainkilo.Text, "kg", vbTextCompare)
        If contain = 0 Then
        Me.txtmainkilo.Text = ""
        End If
     ' Me.txtkilo.Text = CStr(Weight)
     Me.tcommmid.Enabled = True
     Me.Timer7.Enabled = True
     Me.Timer8.Enabled = True
    End If
End Select

End Sub

Private Sub Timer7_Timer()
lblstatus.Caption = statusKilo
If statusKilo = "" Then
lblstatus.Caption = "Unstable"
End If
End Sub

Private Sub Timer8_Timer()
If txtkilo.Tag <> txtkilo.Text Then
statusKilo = "Unstable"
txtkilo.Tag = txtkilo.Text
Else
statusKilo = "Stable"
End If
End Sub

Private Sub tmaincomm_Timer()
Dim symboln As Long
Dim kilo As String
Dim mainKilo As String

If (IsNumeric(Me.txtmainkilo.Text)) Then
current = Me.txtmainkilo.Text
End If

    If Trim(Me.txtmainkilo.Text) = "" Then
             Me.txtmainkilo.Text = current
    Else
     Me.txtmidkilo.Text = ReturnNonAlpha(txtmainkilo.Text)
     symboln = InStr(1, txtmainkilo.Text, commnegative, vbTextCompare)
     If symboln > 0 Then
     Me.txtmidkilo.Text = ReturnNonAlpha(Me.txtmidkilo.Text)
     symbolize = "-"
     Else
     Me.txtmidkilo.Text = ReturnNonAlpha(Me.txtmidkilo.Text)
     symbolize = "+"
     End If
     End If

End Sub
Private Sub tcommmid_Timer()
If txtmidkilo.Text <> "" Then
If txtkilo.Text <> txtmidkilo.Text Then
    txtkilo.Text = txtmidkilo.Text

End If
End If

If lblsymbol.Caption <> symbolize Then
    lblsymbol.Caption = symbolize
End If
End Sub

Private Sub cmbcustomer_Change()
Me.Frame2.height = 3000
productKey = 0
customerKey = 13
If savenum = 1 Then
If Trim$(Me.cmbcustomer.Text) = "" Then
    Set rscustomer = Nothing
    rscustomer.Open "select * from tblcustomer order by customerid ", con, adOpenDynamic, adLockOptimistic
     Me.Frame2.Visible = False
    Set DataGrid1.DataSource = rscustomer
      Set rscustomer = Nothing
    rscustomer.Open "select * from tblcustomer order by customerid ", con, adOpenDynamic, adLockOptimistic
    ElseIf Trim$(Me.cmbcustomer.Text) <> "" Then
    Set rscustomer = Nothing
    rscustomer.Open "select customer_name from tblcustomer where customer_name like'%" & Me.cmbcustomer.Text & "%'", con, adOpenDynamic, adLockOptimistic
     Me.Frame2.Visible = True
     With DataGrid1
    Set .DataSource = rscustomer
    .Columns.Item(0).width = 5000
    End With
    End If
End If
End Sub



Private Sub Timer3_Timer()
Me.lblout.Visible = False
End Sub

Private Sub Timer4_Timer()
Me.lblout.Visible = True
End Sub



Private Sub Timer5_Timer()
Me.Shape1.Visible = True
End Sub

Private Sub Timer6_Timer()
Me.Shape1.Visible = False
End Sub


Private Sub ToggleButton1_Change()
If Me.ToggleButton1.Value = True Then
Me.txtkilo.Locked = True
Me.tmaincomm.Enabled = True
Call addNewLog(currentuser, "Scale Offline")
ElseIf Me.ToggleButton1.Value = False Then
Me.txtkilo.Locked = False
Me.tmaincomm.Enabled = False
End If
End Sub

Private Sub ToggleButton1_Click()
If currentposition = "Weigher" Then
Me.txtkilo.Text = ""
frmscaleoffline.Show 1
End If
End Sub

Private Sub txtkilo_Change()
If CStr(Me.txtkilo.Text) = "" Then
Exit Sub
Else
txtkilo.Text = Int(txtkilo.Text)
End If
If Val(Me.txtkilo.Text) <= 100 Then
Me.Shape1.BackColor = &HC000&
Me.Label21.Caption = "GO!"
Else
Me.Shape1.BackColor = vbRed
Me.Label21.Caption = "STOP!"
End If
End Sub

Private Sub txtkilo_KeyPress(KeyAscii As Integer)
Call numonly(KeyAscii)
If Me.ToggleButton1.Value = False Then
Call numonly(KeyAscii)
End If
End Sub

Private Sub tdatetime_Timer()
Me.lbldate.Caption = Format(Now, "yyyy/MM/dd")
Me.lbltime.Caption = Format(Time, "hh:mm:ss AM/PM")
End Sub

Private Sub Timer1_Timer()
Me.lblin.Visible = True
End Sub

Private Sub Timer2_Timer()
Me.lblin.Visible = False
End Sub
Private Sub enabledall()
Me.cmbcustomer.Enabled = True
Me.cmbproduct.Enabled = True
Me.cmbunit.Enabled = True
Me.cmbdest.Enabled = True
Me.txtqty.Enabled = True
Me.txtprice.Enabled = True
Me.cmdac.Enabled = True
Me.cmdap.Enabled = True
Me.cmdau.Enabled = True
Me.txtremarks.Enabled = True
Me.txtscaleprice.Enabled = True
End Sub
Private Sub disabled()
Me.cmbcustomer.Enabled = True
Me.cmbproduct.Enabled = True
Me.cmbunit.Enabled = True
Me.cmbdest.Enabled = True
Me.txtqty.Enabled = True
Me.txtprice.Enabled = True
Me.cmdac.Enabled = True
Me.cmdap.Enabled = True
Me.cmdau.Enabled = True
Me.txtremarks.Enabled = True
savenum = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
   Call cmdin_Click
    ElseIf KeyCode = vbKeyF2 Then
   Call cmdout_Click
    ElseIf KeyCode = vbKeyF3 Then
    Call cmdsave_Click
     ElseIf KeyCode = vbKeyF9 Then
   Call cmdprint_Click
     ElseIf KeyCode = vbKeyF5 Then
   Call cmdviewin_Click
     ElseIf KeyCode = vbKeyF6 Then
   Call cmdviewall_Click
     ElseIf KeyCode = 27 Then
   Call cmdclose_Click
     ElseIf KeyCode = vbKeyDown Then
     If customerKey = 13 Then
'     Me.Frame2.Visible = True
'     DataGrid1.SetFocus
     Else
     Me.Frame4.Visible = True
     DataGrid2.SetFocus
     End If

  End If
'  If Me.Text1.Text = "1" Then
'  If KeyCode = 13 Then
'  Call cmdsave_Click
'  End If
'  End If
  
  
End Sub


Private Sub txtprice_Change()
Me.txttotalprice.Text = Val(Me.lblnw.Caption) * Val(Me.txtprice.Text)
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
Call numonly(KeyAscii)
End Sub

Private Sub txtqty_Change()
If Val(Me.txtqty.Text) > 0 Then
lblavg.Caption = Val(lblnw.Caption) / txtqty.Text
Me.lblavg.Caption = Format(Me.lblavg.Caption, "###,###,####,#.00")
Else
Me.lblavg.Caption = "0"
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call numonly(KeyAscii)

End Sub

Private Sub txttotalprice_Change()
Me.txttotalprice.Text = Format(Me.txttotalprice.Text, "###,###,##0.00")
End Sub

Public Function ReturnNonAlpha(ByVal sString As String) As String
   Dim i As Integer
   For i = 1 To Len(sString)
       If Mid(sString, i, 1) Like "[0-9]" Then
           ReturnNonAlpha = ReturnNonAlpha & Mid(sString, i, 1)
       End If
   Next i
End Function



