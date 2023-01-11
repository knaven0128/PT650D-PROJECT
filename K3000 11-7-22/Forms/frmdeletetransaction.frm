VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdeletetransaction 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete & Edit Transaction"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   15480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
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
      Height          =   6495
      Left            =   0
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   15495
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
         Left            =   2880
         TabIndex        =   35
         Text            =   "Select Customer........."
         Top             =   5040
         Width           =   3615
      End
      Begin VB.ComboBox cmbproduct 
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
         Left            =   2880
         TabIndex        =   34
         Text            =   "Select Product..........."
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox txtdatetimewo 
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
         TabIndex        =   33
         Tag             =   "txtcom"
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox txtweighout 
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
         TabIndex        =   32
         Tag             =   "txtcom"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtweighin 
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
         TabIndex        =   31
         Tag             =   "txtcom"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtencoder 
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
         TabIndex        =   30
         Tag             =   "txtcom"
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtplatenum 
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
         TabIndex        =   29
         Tag             =   "txtcom"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtnet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   28
         Tag             =   "txtcom"
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtdatetimewi 
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
         TabIndex        =   27
         Tag             =   "txtcom"
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
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
         Left            =   2880
         TabIndex        =   26
         Text            =   "0"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.ComboBox cmbunit 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmdeletetransaction.frx":0000
         Left            =   4680
         List            =   "frmdeletetransaction.frx":0002
         TabIndex        =   25
         Text            =   "cmbproduct"
         Top             =   4440
         Width           =   1695
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
         Left            =   11040
         TabIndex        =   24
         Top             =   5640
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
         Left            =   12960
         TabIndex        =   23
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT NAME:"
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
         Index           =   7
         Left            =   1200
         TabIndex        =   45
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Index           =   6
         Left            =   960
         TabIndex        =   44
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE/TIME WEIGH OUT:"
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
         Index           =   5
         Left            =   480
         TabIndex        =   43
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE/TIME WEIGH IN:"
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
         Index           =   4
         Left            =   600
         TabIndex        =   42
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "NET WEIGHT:"
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
         Index           =   3
         Left            =   1560
         TabIndex        =   41
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "WEIGH OUT:"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   40
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "WEIGH IN:"
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
         TabIndex        =   39
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ENCODER NAME:"
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
         TabIndex        =   38
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PLATE NUMBER:"
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
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF:"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   36
         Top             =   4560
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   9120
      TabIndex        =   16
      Top             =   1850
      Width           =   1215
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
      Left            =   10200
      TabIndex        =   21
      Top             =   7320
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
      Left            =   12120
      TabIndex        =   20
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdeditout 
      Caption         =   "EDIT OUT"
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
      Left            =   8160
      TabIndex        =   19
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdeditin 
      Caption         =   "EDIT IN"
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
      Left            =   6240
      TabIndex        =   18
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdprintout 
      Caption         =   "RE PRINT &OUT"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdprintin 
      Caption         =   "RE PRINT &IN"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5640
      Top             =   1920
   End
   Begin VB.ComboBox cmbcustomer 
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
      ItemData        =   "frmdeletetransaction.frx":0004
      Left            =   2280
      List            =   "frmdeletetransaction.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   6120
      TabIndex        =   1
      Top             =   1850
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   154796035
      CurrentDate     =   42914
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7435
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Tag             =   "txtcom"
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   120
      Picture         =   "frmdeletetransaction.frx":006F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
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
      Left            =   12000
      TabIndex        =   14
      Top             =   2160
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
      Left            =   12720
      TabIndex        =   13
      Top             =   2160
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
      Left            =   12000
      TabIndex        =   12
      Top             =   1800
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
      Left            =   12720
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SORT BY:"
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
      Left            =   720
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
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
      TabIndex        =   9
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE AND EDIT TRANSACTION"
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
      TabIndex        =   8
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes *Once you delete a transaction there's no way to undo!"
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
      Width           =   6090
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click on data before deleting transaction!"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   4155
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Click on data before editing transaction!"
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
      Index           =   8
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   4035
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
Attribute VB_Name = "frmdeletetransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim addnum As Byte
Dim editnum As Byte
Dim printnum As Byte


Private Sub cmbcustomer_Change()
Call cmbcustomer_Click
End Sub

Private Sub cmbcustomer_Click()
On Error GoTo Fixexport
If Me.cmbcustomer.Text = "..Sort By........" Then
Call Form_Load
ElseIf Me.cmbcustomer.Text = "ALL" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing order by weighid ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
ElseIf Me.cmbcustomer.Text = "DAILY" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN like '" & Me.Label1.Caption & "%' order by weighid ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
ElseIf Me.cmbcustomer.Text = "CUSTOMER NAME" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN order by customer_name ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
ElseIf Me.cmbcustomer.Text = "ENCODER" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN order by encoder ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
ElseIf Me.cmbcustomer.Text = "PRODUCT" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN order by product_name ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
ElseIf Me.cmbcustomer.Text = "PLATE NUMBER" Then
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN order by plate_number ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
End If
Fixexport:
If Err.Number = 6160 Then
MsgBox "Check if the excel file is open, Close and export again!", vbOKOnly, "Error!"
End If
End Sub

Private Sub cmdcancel_Click()
With rstruck
        If MsgBox("  Are You Sure You Want To Cancel These Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
         .CancelBatch
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
Load MainForm
End Sub
Private Sub cmddelete_Click()
If Me.cmbcustomer.Text = "" Or Me.cmbcustomer.Text = "..Sort By........" Then
 MsgBox "   Select data first.   ", vbOKOnly, "Warning!"
Else
 If Me.cmbcustomer.Text <> "" And Me.Text1.Text = "" Then
 MsgBox "   Select data you want to delete.   ", vbOKOnly, "Warning!"
 Else
On Error GoTo fixDel
With rstruck
        If .RecordCount > 0 Then
            If MsgBox("  Are You Sure You Want To Delete These Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
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
fixDel:
If Err.Number = 6160 Then
End If
End If
End If
End Sub


Private Sub cmdeditin_Click()
If Me.cmbcustomer.Text = "" Or Me.cmbcustomer.Text = "..Sort By........" Then
 MsgBox "   Select data first.   ", vbOKOnly, "Warning!"
Else
 If Me.cmbcustomer.Text <> "" And Me.Text1.Text = "" Then
 MsgBox "   Select data you want to edit.   ", vbOKOnly, "Warning!"
 Else
editnum = 1
On Error GoTo fixEdit
Me.Frame1.Visible = True
Me.Combo1.Text = rstruck![customer_name]
Me.cmbproduct.Text = rstruck![product_name]
Me.txtplatenum.Text = rstruck![plate_number]
Me.txtencoder.Text = rstruck![encoder]
Me.txtweighin.Text = rstruck![weigh_in]
Me.txtweighout.Text = rstruck![weigh_out]
Me.txtnet.Text = rstruck![Net_weight]
Me.txtdatetimewi.Text = rstruck![DateTime_weighin]
Me.txtdatetimewo.Text = rstruck![DateTime_weighout]
Me.txtweighout.Enabled = False
Me.txtnet.Enabled = False
Me.txtdatetimewo.Enabled = False
Me.txtweighin.Enabled = True
Me.txtnet.Enabled = True
Me.txtdatetimewi.Enabled = True
Me.txtnet.Enabled = False
Me.txtqty.Text = rstruck!qty
Me.cmbunit.Text = rstruck!unit
fixEdit:
If Err.Number = 6061 Then
End If
End If
End If
End Sub





Private Sub cmdprintin_Click()
If Me.cmbcustomer.Text = "" Or Me.cmbcustomer.Text = "..Sort By........" Then
 MsgBox "   Select data first.   ", vbOKOnly, "Warning!"
 End If
 If Me.cmbcustomer.Text <> "" And Me.Text1.Text = "" Then
 MsgBox "   Select data you want to print.   ", vbOKOnly, "Warning!"
 Else
With RTruckScaleIN
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where consec_no like '" & Me.Text1.Text & "'"
                .Restart
                .Show 1
                '.PrintReport False
            End With
End If
End Sub

Private Sub cmdprintout_Click()
If Me.cmbcustomer.Text = "" Or Me.cmbcustomer.Text = "..Sort By........" Then
 MsgBox "   Select data first.   ", vbOKOnly, "Warning!"
 End If
 If Me.cmbcustomer.Text <> "" And Me.Text1.Text = "" Then
 MsgBox "   Select data you want to print.   ", vbOKOnly, "Warning!"
 Else
With RTruckScaleOUT
                .ado2.Connection = con
                .ado2.Source = "select * from tblweighing where consec_no like '" & Me.Text1.Text & "'"
                .Restart
                .Show 1
                '.PrintReport False
            End With
            End If
End Sub

Private Sub cmdsave_Click()
Select Case editnum
Case 1
With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                "set plate_number='" & Me.txtplatenum.Text & "'," & _
                 "encoder='" & Me.txtencoder.Text & "'," & _
                 "weigh_in='" & Me.txtweighin.Text & "'," & _
                 "datetime_weighin='" & Me.txtdatetimewi.Text & "'," & _
                 "customer_name= '" & Me.Combo1.Text & "'," & _
                 "QTY='" & Val(Me.txtqty.Text) & "'," & _
                "UNIT='" & Me.cmbunit.Text & "'," & _
                 "product_name='" & Me.cmbproduct.Text & "'" & _
                " where weighid=" & rstruck!weighid
              .Execute
            End With
            Call Emptyctl(Me, "txtcom")
               
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
        Call cmbcustomer_Click
    Me.Frame1.Visible = False
    Case 2
    With cmd
       .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                "set plate_number='" & Me.txtplatenum.Text & "'," & _
                 "encoder='" & Me.txtencoder.Text & "'," & _
                 "weigh_in='" & Me.txtweighin.Text & "'," & _
                 "weigh_out='" & Me.txtweighout.Text & "'," & _
                 "net_weight='" & Me.txtnet.Text & "'," & _
                 "QTY='" & Val(Me.txtqty.Text) & "'," & _
                "UNIT='" & Me.cmbunit.Text & "'," & _
                 "datetime_weighin='" & Me.txtdatetimewi.Text & "'," & _
                 "datetime_weighout='" & Me.txtdatetimewo.Text & "'," & _
                 "customer_name= '" & Me.Combo1.Text & "'," & _
                 "product_name='" & Me.cmbproduct.Text & "'" & _
                " where weighid=" & rstruck!weighid
              .Execute
            End With
            Call Emptyctl(Me, "txtcom")
               
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
        Call cmbcustomer_Click
    Me.Frame1.Visible = False
    End Select
End Sub

Private Sub DataGrid1_Click()
Me.Text1.Text = rstruck![consec_no]
End Sub

Private Sub Form_Load()
Me.Timer1.Enabled = True
sDateString = Format(Now, "m/d/yyyy")
Call cand
Call cand1
Call cand2
End Sub

Private Sub cmdsearch_Click()
Set rstruck = Nothing
rstruck.Open "Select * from tblweighing where DateTime_WeighIN like '" & Me.DTPicker1.Value & "%'order by plate_number ", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rstruck
End Sub

Private Sub cmdeditout_Click()
If Me.cmbcustomer.Text = "" Or Me.cmbcustomer.Text = "..Sort By........" Then
 MsgBox "   Select data first.   ", vbOKOnly, "Warning!"
Else
 If Me.cmbcustomer.Text <> "" And Me.Text1.Text = "" Then
 MsgBox "   Select data you want to edit.   ", vbOKOnly, "Warning!"
 Else
editnum = 2
Me.Frame1.Visible = True
Me.txtplatenum.Text = rstruck![plate_number]
Me.txtencoder.Text = rstruck![encoder]
Me.txtweighin.Text = rstruck![weigh_in]
Me.txtweighout.Text = rstruck![weigh_out]
Me.txtnet.Text = rstruck![Net_weight]
Me.txtdatetimewi.Text = rstruck![DateTime_weighin]
Me.txtdatetimewo.Text = rstruck![DateTime_weighout]
Me.Combo1.Text = rstruck![customer_name]
Me.cmbproduct.Text = rstruck![product_name]
Me.txtqty.Text = rstruck![qty]
 Me.cmbunit.Text = rstruck![unit]
Me.txtweighin.Enabled = False
Me.txtnet.Enabled = False
Me.txtdatetimewi.Enabled = False
Me.txtweighout.Enabled = True
Me.txtnet.Enabled = True
Me.txtdatetimewo.Enabled = True
Me.txtnet.Enabled = False
End If
End If
End Sub



Private Sub Timer1_Timer()
Me.Label9.Caption = Format(Now, "mmmm dd, yyyy")
Me.Label11.Caption = Format(Now, "hh:mm:ss AM/PM")
Me.Label1.Caption = sDateString
End Sub
Private Sub datasize()
With DataGrid1
            .Columns.Item(0).Visible = False
            .Columns.Item(1).width = 1500
            .Columns.Item(2).width = 1500
            .Columns.Item(3).width = 1000
            .Columns.Item(4).width = 1000
            .Columns.Item(5).width = 1000
            .Columns.Item(6).width = 1800
            .Columns.Item(7).width = 2100
            .Columns.Item(8).width = 2100
            .Columns.Item(9).width = 1800
            .Columns.Item(10).width = 1800
            .Columns.Item(11).Visible = False
            .Columns.Item(12).Visible = False
End With
End Sub
Private Sub cand()
Set rscustomer = Nothing
rscustomer.Open "Select*from tblcustomer ", con, 3, 3

    With rscustomer
        .Filter = 0
        While Not .EOF
            Me.Combo1.AddItem .Fields("Customer_name")
            .MoveNext
        Wend
    End With
rscustomer.Close
End Sub
Private Sub cand1()
Set rsproduct = Nothing
rsproduct.Open "Select*from tblproduct ", con, 3, 3

    With rsproduct
        .Filter = 0
        While Not .EOF
            Me.cmbproduct.AddItem .Fields("Product_name")
            .MoveNext
        Wend
    End With
rsproduct.Close
End Sub
Private Sub cand2()
Set rsum = Nothing
rsum.Open "Select*from tblunitmeasure ", con, 3, 3

    With rsum
        .Filter = 0
        While Not .EOF
            Me.cmbunit.AddItem .Fields("Unit_Symbol")
            .MoveNext
        Wend
    End With
rsum.Close
End Sub

Private Sub txtnet_Change()
Me.txtnet.Text = Str(Abs(Val(txtnet.Text)))
End Sub

Private Sub txtweighin_Change()
Me.txtnet.Text = Val(Me.txtweighin.Text) - Val(Me.txtweighout.Text)
End Sub

Private Sub txtweighout_Change()
Me.txtnet.Text = Val(Me.txtweighin.Text) - Val(Me.txtweighout.Text)
End Sub


