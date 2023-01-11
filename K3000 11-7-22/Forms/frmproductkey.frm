VERSION 5.00
Begin VB.Form frmproductkey 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weighing Scale Product Registration"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin Weighing.jcbutton cmdregister 
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Caption         =   "&Register"
      Picture         =   "frmproductkey.frx":0000
      UseMaskCOlor    =   -1  'True
      MaskColor       =   1
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   480
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   3720
      Width           =   4935
   End
   Begin Weighing.jcbutton cmdclose 
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Caption         =   "&Close"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "frmproductkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

