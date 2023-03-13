VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlist 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Records"
   ClientHeight    =   9465
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   16065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   16065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "DATA ENTRY"
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
      Height          =   9495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   16095
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
         ItemData        =   "frmlist.frx":0000
         Left            =   2760
         List            =   "frmlist.frx":0002
         TabIndex        =   53
         Top             =   8040
         Width           =   4215
      End
      Begin VB.TextBox Combo1 
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
         TabIndex        =   46
         Tag             =   "txtcom"
         Top             =   6960
         Width           =   3615
      End
      Begin VB.TextBox txtavg 
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
         TabIndex        =   43
         Tag             =   "txtcom"
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox txtscaleprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   41
         Text            =   "0"
         Top             =   6480
         Width           =   3495
      End
      Begin VB.TextBox txttotalprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   39
         Text            =   "0"
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox txtprice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   37
         Text            =   "0"
         Top             =   5520
         Width           =   3495
      End
      Begin VB.TextBox txtremarks 
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
         Height          =   810
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   8520
         Width           =   5535
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
         Left            =   13200
         TabIndex        =   15
         Top             =   8760
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
         Left            =   11280
         TabIndex        =   14
         Top             =   8760
         Width           =   1815
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
         ItemData        =   "frmlist.frx":0004
         Left            =   4680
         List            =   "frmlist.frx":0006
         TabIndex        =   13
         Text            =   "Select Unit................"
         Top             =   5040
         Width           =   1695
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
         TabIndex        =   12
         Text            =   "0"
         Top             =   5040
         Width           =   1695
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
         TabIndex        =   11
         Tag             =   "txtcom"
         Top             =   3840
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
         TabIndex        =   10
         Tag             =   "txtcom"
         Top             =   2640
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
         TabIndex        =   9
         Tag             =   "txtcom"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtweigher 
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
         TabIndex        =   8
         Tag             =   "txtcom"
         Top             =   840
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
         TabIndex        =   7
         Tag             =   "txtcom"
         Top             =   1440
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
         TabIndex        =   6
         Tag             =   "txtcom"
         Top             =   2040
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
         TabIndex        =   5
         Tag             =   "txtcom"
         Top             =   4440
         Width           =   3615
      End
      Begin VB.ComboBox cmbproduct 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         TabIndex        =   4
         Text            =   "Select Product..........."
         Top             =   7440
         Width           =   3615
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
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   54
         Top             =   8040
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "AVERAGE:"
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
         Index           =   1
         Left            =   1680
         TabIndex        =   44
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "T.S PRICE:"
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
         Left            =   1440
         TabIndex        =   42
         Top             =   6600
         Width           =   1095
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
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   40
         Top             =   6120
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
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   38
         Top             =   5640
         Width           =   735
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
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   36
         Top             =   8640
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY:"
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
         TabIndex        =   25
         Top             =   5160
         Width           =   1215
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
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATOR NAME:"
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
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   1695
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
         TabIndex        =   22
         Top             =   1560
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
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
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
         TabIndex        =   20
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   19
         Top             =   3960
         Width           =   2175
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
         TabIndex        =   18
         Top             =   4560
         Width           =   2295
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
         TabIndex        =   17
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Index           =   7
         Left            =   1200
         TabIndex        =   16
         Top             =   7440
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdprintall 
      Caption         =   "RE PRINT &ALL"
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
      Left            =   4680
      TabIndex        =   52
      Top             =   8040
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Weigh ID"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Weighing No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Plate Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Transaction Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Gross"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tare"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Net Weight"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Average"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Commodity"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Destination"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Price"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Total Price"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Quantity"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Unit"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Date/Time Weigh In "
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Date/Time Weigh Out"
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Scale Price"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Weigher"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Remarks"
         Object.Width           =   5115
      EndProperty
   End
   Begin VB.CommandButton cmdLoadData 
      Caption         =   "LOAD DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Picture         =   "frmlist.frx":0008
      TabIndex        =   50
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtAccess 
      Height          =   375
      Left            =   -120
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      Picture         =   "frmlist.frx":366B
      TabIndex        =   33
      Top             =   960
      Width           =   1335
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
      Left            =   120
      TabIndex        =   32
      Tag             =   "txtcom"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   720
      TabIndex        =   31
      Top             =   8040
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
      Left            =   2640
      TabIndex        =   30
      Top             =   8040
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
      Left            =   6600
      TabIndex        =   29
      Top             =   8040
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
      Left            =   8520
      TabIndex        =   28
      Top             =   8040
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
      Left            =   12360
      TabIndex        =   27
      Top             =   8040
      Width           =   1815
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
      Left            =   10440
      TabIndex        =   26
      Top             =   8040
      Width           =   1815
   End
   Begin VB.ComboBox cboFields 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmlist.frx":6CCE
      Left            =   240
      List            =   "frmlist.frx":6CD0
      TabIndex        =   2
      Top             =   960
      Width           =   2295
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
      Left            =   2640
      TabIndex        =   1
      Tag             =   "txtcom"
      Top             =   960
      Width           =   6975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   8160
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   47
      Top             =   240
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
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   130088963
      CurrentDate     =   42922
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4560
      TabIndex        =   48
      Top             =   240
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
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   130088963
      CurrentDate     =   42922
   End
   Begin VB.Label Label7 
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
      Left            =   3600
      TabIndex        =   51
      Top             =   360
      Width           =   1335
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
      Left            =   240
      TabIndex        =   49
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Items Count: 0"
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
      Left            =   360
      TabIndex        =   34
      Top             =   7200
      Width           =   4815
   End
End
Attribute VB_Name = "frmlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim list_item As ListItem
Dim lst1 As ListItem
Dim X As New Class1
Dim i As Integer
Dim editnum As Integer
Public weighStatus As String
Dim update_data As String
Dim update_dataOut As String
Dim plateNumber As String
Dim strDate As String
Private Sub cmbproduct_DropDown()
Me.cmbproduct.Clear
cmbproduct.AddItem "NA"
Set rsproduct = Nothing
With rsproduct
.Open "Select * from tblproduct", con, 3, 3
Do Until .EOF
cmbproduct.AddItem !product_name
.MoveNext
Loop
End With
rsproduct.Close
End Sub

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




Private Sub cmdcancel_Click()
With rstruck
        If MsgBox("  Are You Sure You Want To Cancel This Record?   ", vbQuestion + vbYesNo, "Prompt") = vbYes Then
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
Me.Timer2.Enabled = True
End Sub


Private Sub cmddelete_Click()
On Error GoTo fixDel
If currentposition = "Weigher" And Me.txtAccess.Text = "" Then
frmReprint.Show 1
Else
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
  If MsgBox("Are you sure to delete this record?", vbQuestion + vbYesNo, "Prompt") = vbYes Then
            With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                " set delstatus=" & 1 & _
                " where consec_no=" & Me.Text1.Text
              .Execute
              Call addNewLog(currentuser, "Delete Weighing Data - " + "transaction no: " + Me.Text1.Text + " - plate number: " + plateNumber)
            End With
             Call listing
             MsgBox "Record successfully deleted.   ", vbOKOnly, "Success!"
            Else
            Cancel = 1
            End If
End If
End If
fixDel:
If Err.Number = 6160 Then
End If
End Sub

Private Sub cmdeditin_Click()
On Error GoTo fixEdit
editnum = 1
If currentposition = "Weigher" And Me.txtAccess.Text = "" Then
frmReprint.Show 1
Else
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
Me.Frame1.Visible = True
Me.txtweighout.Enabled = False
Me.txtnet.Enabled = False
Me.txtdatetimewo.Enabled = False
Me.txtqty.Enabled = False
Me.txtavg.Enabled = False
Me.txtscaleprice.Enabled = False
Me.txtweighin.Enabled = True
Me.txtnet.Enabled = False
Me.txtdatetimewi.Enabled = True
Me.txtnet.Enabled = False
Set rstruck = Nothing
rstruck.Open "select * from tblweighing where consec_no = '" & Me.Text1.Text & "'", con, 3, 3
Me.Combo1.Text = rstruck![customer_name]
Me.txtdatetimewi.Text = rstruck![datetime_weighin]
Me.cmbproduct.Text = rstruck![product_name]
Me.txtplatenum.Text = rstruck![plate_number]
Me.txtweigher.Text = rstruck![weigher]
Me.txtweighin.Text = rstruck![weigh_in]
Me.cmbunit.Text = rstruck![unit]
Me.txtqty.Text = rstruck![qty]
Me.txtprice.Text = rstruck![Price]
Me.txttotalprice.Text = rstruck![totalprice]
Me.txtavg.Text = rstruck![Average]
Me.txtscaleprice.Text = rstruck![scale_price]
Me.cmbdest.Text = rstruck![Destination] & vbNullString

Me.txtremarks.Text = rstruck![Remarks]
End If
End If
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdeditout_Click()
On Error GoTo fixEdit
editnum = 2
If currentposition = "Weigher" And Me.txtAccess.Text = "" Then
frmReprint.Show 1
Else
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else

Set rstruck = Nothing
rstruck.Open "select * from tblweighing where consec_no = '" & Me.Text1.Text & "'", con, 3, 3

Me.Frame1.Visible = True
Me.txtplatenum.Text = rstruck![plate_number]
Me.txtweigher.Text = rstruck![weigher]
Me.txtweighin.Text = rstruck![weigh_in]
Me.txtweighout.Text = rstruck![weigh_out]
Me.txtnet.Text = rstruck![net_weight]
Me.txtdatetimewi.Text = rstruck![datetime_weighin]
Me.txtdatetimewo.Text = rstruck![datetime_weighout]
Me.Combo1.Text = rstruck![customer_name]
Me.cmbproduct.Text = rstruck![product_name]
Me.txtqty.Text = rstruck![qty]
Me.cmbunit.Text = rstruck![unit]
Me.cmbdest.Text = rstruck![Destination] & vbNullString
Me.txtavg.Text = rstruck![Average]
Me.txtscaleprice.Text = rstruck![scale_price]
Me.txtremarks.Text = rstruck![Remarks]
Me.txtweighin.Enabled = False
Me.txtnet.Enabled = False
Me.txtdatetimewi.Enabled = False
Me.txtweighout.Enabled = True
Me.txtnet.Enabled = True
Me.txtscaleprice.Enabled = True
Me.txtprice.Enabled = True
Me.txttotalprice.Enabled = False
Me.txtdatetimewo.Enabled = True
Me.txtnet.Enabled = False
End If
End If
fixEdit:
If Err.Number = 6061 Then
End If
End Sub

Private Sub cmdLoadData_Click()
Call listing
End Sub

Private Sub cmdprintall_Click()
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
Call addNewLog(currentuser, "Reprint OUT - Transaction No.: " + Me.Text1.Text)
                With PrintAllScale
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where consec_no ='" & Me.Text1.Text & "'"
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
                '.Label14.Visible = False
                .Show 1
            End With
End If
End If
End Sub

Private Sub cmdprintin_Click()
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
'With RTruckScaleIN
Call addNewLog(currentuser, "Reprint IN - Transaction No.: " + Me.Text1.Text)
With OliverScaleIN
         .ado1.Connection = con
               .ado1.Source = "select * from tblweighing where consec_no ='" & Me.Text1.Text & "'"
              .Restart
                .Label17.Caption = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
'               .Field7.Visible = False
'               .Field9.Visible = False
'               .Field10.Visible = False
'               .Field11.Visible = False
''               .Field12.Visible = False
''               .Field13.Visible = False
'               .Field14.Visible = False
'               .Label29.Visible = False
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
                End With

End If
End Sub

Private Sub cmdprintout_Click()
Me.txtAccess.Text = ""
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
If Me.Text1.Text = "" Then
MsgBox "Please Select Data!" & vbNewLine & "Please Contact the Programmer.", vbCritical, "Invalid"
Else
Call addNewLog(currentuser, "Reprint OUT - Transaction No.: " + Me.Text1.Text)
                With OliverScaleOUT
                .ado1.Connection = con
                .ado1.Source = "select * from tblweighing where consec_no ='" & Me.Text1.Text & "'"
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
                '.Label14.Visible = False
                .Show 1
            End With
End If
End If
End Sub

Private Sub cmdrecomp_Click()
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

Private Sub cmdrefresh_Click()
Call listing
End Sub

Private Sub cmdsave_Click()
Select Case editnum
Case 1
update_data = dataIn_Updated
With cmd
                .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                "set plate_number='" & Me.txtplatenum.Text & "'," & _
                 "weigher='" & Me.txtweigher.Text & "'," & _
                 "weigh_in='" & Me.txtweighin.Text & "'," & _
                "net_weight='" & Me.txtnet.Text & "'," & _
                 "datetime_weighin='" & Format$(Me.txtdatetimewi.Text, "yyyy/mm/dd hh:mm:ss") & "'," & _
                 "customer_name= '" & Me.Combo1.Text & "'," & _
                 "QTY='" & Val(Me.txtqty.Text) & "'," & _
                "UNIT='" & Me.cmbunit.Text & "'," & _
                "Destination='" & Me.cmbdest.Text & "'," & _
                "Average='" & Me.txtavg.Text & "'," & _
                 "product_name='" & Me.cmbproduct.Text & "'," & _
                  "remarks='" & Me.txtremarks.Text & "'" & _
                " where weighid=" & rstruck!weighid
              .Execute
              Call addNewLog(currentuser, "Update Weigh In Data - " + "transacion no.:" + Me.Text1.Text + update_data)
            End With
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
         Call Emptyctl(Me, "txtcom")
        Call listing
    Me.Frame1.Visible = False
Case 2
   update_dataOut = dataOut_Updated
If Val(Me.txtqty.Text) = 0 Then
  Me.txttotalprice.Text = Val(Me.txtnet.Text) * Val(Me.txtprice.Text)
    Me.txtavg.Text = 0
    Else
      Me.txttotalprice.Text = Val(Me.txtnet.Text) * Val(Me.txtprice.Text)
    Me.txtavg.Text = Val(Me.txtnet.Text) / Val(Me.txtqty.Text)
     End If

    With cmd
       .ActiveConnection = con
                .CommandType = adCmdText
                .CommandText = "update tblweighing " & _
                "set plate_number='" & Me.txtplatenum.Text & "'," & _
                 "weigher='" & Me.txtweigher.Text & "'," & _
                 "weigh_in='" & Me.txtweighin.Text & "'," & _
                 "weigh_out='" & Me.txtweighout.Text & "'," & _
                 "net_weight='" & Me.txtnet.Text & "'," & _
                 "QTY='" & Val(Me.txtqty.Text) & "'," & _
                "UNIT='" & Me.cmbunit.Text & "'," & _
                "Destination='" & Me.cmbdest.Text & "'," & _
                "Average='" & Me.txtavg.Text & "'," & _
                "price='" & CDbl(Me.txtprice.Text) & "'," & _
                "totalprice='" & CDbl(Me.txttotalprice.Text) & "'," & _
                "scale_price='" & CDbl(Me.txtscaleprice.Text) & "'," & _
                 "datetime_weighout='" & Format$(Me.txtdatetimewo.Text, "yyyy/mm/dd hh:mm:ss") & "'," & _
                 "customer_name= '" & Me.Combo1.Text & "'," & _
                 "product_name='" & Me.cmbproduct.Text & "'," & _
                  "remarks='" & Me.txtremarks.Text & "'" & _
                " where weighid=" & rstruck!weighid
              .Execute
               Call addNewLog(currentuser, "Update Weigh Out Data - " + "transacion no.:" + Me.Text1.Text + update_dataOut)
            End With
        MsgBox "   Record Successfully Updated.   ", vbOKOnly, "Success!"
           Call Emptyctl(Me, "txtcom")
        Call listing
    Me.Frame1.Visible = False
    End Select

End Sub



Private Sub Form_Load()
strDate = Format(Now, "yyyy/MM/dd")
Me.DTPicker1.Value = DateSerial(Year(strDate), Month(strDate), 1)
Me.DTPicker2.Value = strDate
Me.txtAccess.Text = ""
Me.cboFields.AddItem "Weighing No"
Me.cboFields.AddItem "Plate_Number"
Me.cboFields.AddItem "Operator"
Me.cboFields.AddItem "Customer_Name"
Me.cboFields.AddItem "Commodity"
If listnum = 1 Then
weighStatus = "IN"
Me.DTPicker1.Value = "01/01/1990"
Me.DTPicker2.Value = strDate
Me.DTPicker1.Visible = False
Me.DTPicker2.Visible = False
Me.Label3.Visible = False
Me.Label7.Visible = False
Me.cmdLoadData.Visible = False
Call listing
Me.cmdprintall.Visible = False
Me.cmdprintout.Visible = False
Me.cmdprintin.Left = 8480
Me.cmddelete.Left = 10400
Me.cmdeditout.Visible = False
ElseIf listnum = 2 Then
Me.DTPicker1.Visible = True
Me.DTPicker2.Visible = True
Me.Label3.Visible = True
Me.Label7.Visible = True
Me.cmdLoadData.Visible = True
Me.cmdprintout.Visible = True
Me.cmdeditout.Visible = True
weighStatus = "OUT"
End If
Me.width = 0
Me.Timer1.Enabled = True
Set rscompany = Nothing
rscompany.Open "select * from tblcompany", ocn, 3, 3
If currentposition = "Weigher" And Me.txtAccess.Text = "" Then
If listnum = 1 Then
Me.cmdprintin.Left = 8400
Me.cmdprintall.Visible = False
Else
Me.cmdprintin.Left = 6480
cmdprintall.Left = 10400
Me.cmdprintall.Visible = True
End If
Me.cmdeditin.Visible = False
Me.cmdeditout.Visible = False
Me.cmddelete.Visible = False
Me.cmdprintout.Left = 8400

End If
If listnum = 2 Then
Set rs = Nothing
rs.Open "Select * from tblweighing where status like 'OUT' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & strDate & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & strDate & "' order by weighid desc ", con, 3, 3
Call listview
Set rs = Nothing
rs.Open "Select COUNT(*) as sCount from tblweighing where status like 'OUT' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & strDate & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & strDate & "' order by weighid desc ", con, 3, 3
Me.Label1.Caption = "Total Items Count: " & rs!sCount
End If
End Sub



Private Sub lvData_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.Text1.Text = lvData.SelectedItem.SubItems(1)
plateNumber = lvData.SelectedItem.SubItems(2)
End Sub

Private Sub mad_Click()
frmpermission.Show 1
End Sub

Private Sub Timer1_Timer()
Me.Left = 800
Me.Top = 1000
Me.width = Me.width + 700
If Me.width >= 16000 Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Me.width = Me.width - 700
If Me.width < 1000 Then
Timer2.Enabled = False
Unload Me
End If
End Sub


Private Sub listing()
Select Case listnum
Case 1
Set rs = Nothing
rs.Open "Select * from tblweighing where status like 'IN' and delstatus = 0 order by weighid desc", con, 3, 3
Call listview
Set rs = Nothing
rs.Open "Select COUNT(*) as sCount from tblweighing where Status like 'IN' and delstatus = 0 order by weighid desc", con, adOpenStatic, adLockOptimistic
Me.Label1.Caption = "Total Items Count: " & rs!sCount
Case 2
Set rs = Nothing
rs.Open "Select * from tblweighing where date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d') >='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d') <='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "'  and delstatus = 0 and status = 'OUT' order by weighid", con, 3, 3
Call listview
Set rs = Nothing
rs.Open "Select COUNT(*) as sCount from tblweighing where date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d') >='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d') <='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' and status = 'OUT' and delstatus = 0 order by weighid desc ", con, 3, 3
Me.Label1.Caption = "Total Items Count: " & rs!sCount
 End Select
End Sub
 '.txtfirstname.Text = lvStudentInfo.SelectedItem.SubItems(2)
Public Sub listview()
'Set lvData.SmallIcons = ImageList1
Me.lvData.ListItems.Clear
Do Until rs.EOF

    Set lst1 = Me.lvData.ListItems.Add(, , rs!weighid)
        With lst1
            .SubItems(1) = rs!consec_no & vbNullString
            .SubItems(2) = rs!plate_number & vbNullString
            .SubItems(3) = rs!customer_name & vbNullString
            .SubItems(4) = Format$(rs!transaction_date, "yyyy/MM/dd") & vbNullString
            If (rs!weigh_in > rs!weigh_out) Then
'            gross
             .SubItems(5) = rs!weigh_in & vbNullString
             .SubItems(6) = rs!weigh_out & vbNullString
            Else
'            tare
            .SubItems(5) = rs!weigh_out & vbNullString
            .SubItems(6) = rs!weigh_in & vbNullString
            End If
    
            .SubItems(7) = rs!net_weight & vbNullString
            .SubItems(8) = rs!Average & vbNullString
            .SubItems(9) = rs!product_name & vbNullString
            .SubItems(10) = rs!Destination & vbNullString
            .SubItems(11) = rs!Price & vbNullString
            .SubItems(12) = rs!totalprice & vbNullString
            .SubItems(13) = rs!qty & vbNullString
            .SubItems(14) = rs!unit & vbNullString
            .SubItems(15) = rs!datetime_weighin & vbNullString
            .SubItems(16) = rs!datetime_weighout & vbNullString
            .SubItems(17) = rs!Status & vbNullString
            .SubItems(18) = rs!scale_price & vbNullString
            .SubItems(19) = rs!weigher & vbNullString
            .SubItems(20) = rs!Remarks & vbNullString
            
        End With
    rs.MoveNext
Loop
'With lvData
'For i = 3 To .ColumnHeaders.Count
'.ColumnHeaders(i).width = 2000
'Next i
''x.SetListViewAlternateColor lvData
'End With
End Sub

Private Sub txtavg_Change()

Me.txtavg.Text = Format(Me.txtavg.Text, "###,###,####,#.00")
End Sub

Private Sub txtnet_Change()
txtnet.Text = Str(Abs(Val(txtnet.Text)))
End Sub

Private Sub txtprice_Change()
On Error GoTo Err:
Me.txttotalprice.Text = Val(Me.txtnet.Text) * Val(Me.txtprice.Text)
Exit Sub
Err:
Exit Sub
End Sub

Private Sub txtscaleprice_Change()
Me.txtscaleprice.Text = Format(Me.txtscaleprice.Text, "###,###,##0.00")
End Sub

Private Sub txtsearch_Change()
If Me.cboFields.Text = "" Then
MsgBox "Select Data to Search!", vbCritical, "Warning"
ElseIf Trim$(Me.txtsearch.Text) = "" Then
Call listing
ElseIf Trim$(Me.txtsearch.Text) <> "" Then
    If Me.cboFields.Text = "Commodity" Then
     Set rs = Nothing
    rs.Open "select * from tblweighing where Product_Name like'%" & Me.txtsearch.Text & "%' AND status like '%" & weighStatus & "%' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid desc", con, adOpenDynamic, adLockOptimistic
       Call listview
    ElseIf Me.cboFields.Text = "Weighing No" Then
     Set rs = Nothing
    rs.Open "select * from tblweighing where consec_no like'%" & Me.txtsearch.Text & "%' AND status like '%" & weighStatus & "%' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid desc", con, adOpenDynamic, adLockOptimistic
       Call listview
    ElseIf Me.cboFields.Text = "Operator" Then
       Set rs = Nothing
    rs.Open "select * from tblweighing where weigher like'%" & Me.txtsearch.Text & "%' AND status like '%" & weighStatus & "%' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid desc", con, adOpenDynamic, adLockOptimistic
       Call listview
    Else
    Set rs = Nothing
    rs.Open "select * from tblweighing where " & Me.cboFields.Text & " like'%" & Me.txtsearch.Text & "%' AND status like '%" & weighStatus & "%' and delstatus = 0 and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')>='" & Format$(DTPicker1.Value, "yyyy/mm/dd") & "' and date_format(cast(`tblweighing`.`datetime_weighout` as date),'%Y/%m/%d')<='" & Format$(DTPicker2.Value, "yyyy/mm/dd") & "' order by weighid desc", con, adOpenDynamic, adLockOptimistic
       Call listview
    End If
       Else
   Call listing
End If
End Sub

Private Sub txttotalprice_Change()
On Error GoTo Err:

Me.txttotalprice.Text = Format(Me.txttotalprice.Text, "###0.00")
Exit Sub
Err:
End Sub

Private Sub txtweighin_Change()
Me.txtnet.Text = Val(Me.txtweighin.Text) - Val(Me.txtweighout.Text)
End Sub

Private Sub txtweighout_Change()
Me.txtnet.Text = Val(Me.txtweighin.Text) - Val(Me.txtweighout.Text)
      Me.txttotalprice.Text = Val(Me.txtnet.Text) * Val(Me.txtprice.Text)
End Sub
Private Function dataIn_Updated() As String

If Me.txtplatenum.Text <> rstruck![plate_number] Then
dataIn_Updated = dataIn_Updated + " - plate number: " + rstruck![plate_number]
End If

If Me.txtweighin.Text <> rstruck![weigh_in] Then
dataIn_Updated = dataIn_Updated + " - weigh in: " + CStr(rstruck![weigh_in])
End If

If Me.txtdatetimewi.Text <> rstruck![datetime_weighin] Then
dataIn_Updated = dataIn_Updated + " - datetime weighin: " + CStr(rstruck![datetime_weighin])
End If

If Me.Combo1.Text <> rstruck![customer_name] Then
dataIn_Updated = dataIn_Updated + " - customer name: " + rstruck![customer_name]
End If

If Me.cmbproduct.Text <> rstruck![product_name] Then
dataIn_Updated = dataIn_Updated + " - commodity: " + rstruck![product_name]
End If

If Me.txtweigher.Text <> rstruck![weigher] Then
dataIn_Updated = dataIn_Updated + " - operator: " + rstruck![weigher]
End If


If Me.cmbunit.Text <> rstruck![unit] Then
dataIn_Updated = dataIn_Updated + " - unit: " + rstruck![unit]
End If

If Me.txtqty.Text <> rstruck![qty] Then
dataIn_Updated = dataIn_Updated + " - quantity: " + CStr(rstruck![qty])
End If

If Me.txtprice.Text <> rstruck![Price] Then
dataIn_Updated = dataIn_Updated + " - price: " + CStr(rstruck![Price])
End If

If Me.txttotalprice.Text <> rstruck![totalprice] Then
dataIn_Updated = dataIn_Updated + " - total price: " + CStr(rstruck![totalprice])
End If

If Me.txtavg.Text <> rstruck![Average] Then
dataIn_Updated = dataIn_Updated + " - average: " + CStr(rstruck![Average])
End If

If Me.txtscaleprice.Text <> rstruck![scale_price] Then
dataIn_Updated = dataIn_Updated + " - scale price: " + CStr(rstruck![scale_price])
End If

If Me.txtremarks.Text <> rstruck![Remarks] Then
dataIn_Updated = dataIn_Updated + " - remarks: " + rstruck![Remarks]
End If

End Function

Private Function dataOut_Updated() As String
If Me.txtplatenum.Text <> rstruck![plate_number] Then
dataOut_Updated = dataOut_Updated + " - plate number: " + rstruck![plate_number]
End If

If Me.txtweigher.Text <> rstruck![weigher] Then
dataOut_Updated = dataOut_Updated + " - operator: " + rstruck![weigher]
End If

If Me.txtweighin.Text <> rstruck![weigh_in] Then
dataOut_Updated = dataOut_Updated + " - weigh in: " + CStr(rstruck![weigh_in])
End If

If Me.txtweighout.Text <> rstruck![weigh_out] Then
dataOut_Updated = dataOut_Updated + " - weigh out: " + CStr(rstruck![weigh_out])
End If

If Me.txtnet.Text <> rstruck![net_weight] Then
dataOut_Updated = dataOut_Updated + " - net weight: " + CStr(rstruck![net_weight])
End If

If Me.txtavg.Text <> rstruck![Average] Then
dataOut_Updated = dataOut_Updated + " - average: " + CStr(rstruck![Average])
End If

If Me.txtdatetimewi.Text <> rstruck![datetime_weighin] Then
dataOut_Updated = dataOut_Updated + " - datetime weighin: " + CStr(rstruck![datetime_weighin])
End If

If Me.txtdatetimewo.Text <> rstruck![datetime_weighout] Then
dataOut_Updated = dataOut_Updated + " - datetime weighout: " + CStr(rstruck![datetime_weighout])
End If

If Me.Combo1.Text <> rstruck![customer_name] Then
dataOut_Updated = dataOut_Updated + " - customer name: " + rstruck![customer_name]
End If

If Me.cmbproduct.Text <> rstruck![product_name] Then
dataOut_Updated = dataOut_Updated + " - commodity: " + rstruck![product_name]
End If

If Me.cmbunit.Text <> rstruck![unit] Then
dataOut_Updated = dataOut_Updated + " - unit: " + rstruck![unit]
End If

If Me.txtqty.Text <> rstruck![qty] Then
dataOut_Updated = dataOut_Updated + " - quantity: " + CStr(rstruck![qty])
End If

If Me.txtprice.Text <> rstruck![Price] Then
dataOut_Updated = dataOut_Updated + " - price: " + CStr(rstruck![Price])
End If

If Me.txttotalprice.Text <> rstruck![totalprice] Then
dataOut_Updated = dataOut_Updated + " - total price: " + CStr(rstruck![totalprice])
End If

If Me.txtscaleprice.Text <> rstruck![scale_price] Then
dataOut_Updated = dataOut_Updated + " - scale price: " + CStr(rstruck![scale_price])
End If

If Me.txtremarks.Text <> rstruck![Remarks] Then
dataOut_Updated = dataOut_Updated + " - remarks: " + rstruck![Remarks]
End If

End Function


