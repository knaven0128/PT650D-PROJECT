VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RTruckScaleIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Truck Scale IN"
   ClientHeight    =   15615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   28560
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   50377
   _ExtentY        =   27543
   SectionData     =   "RTruckScaleIN.dsx":0000
End
Attribute VB_Name = "RTruckScaleIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
If Me.lblemail.Visible = False Then
Me.lblcontact.Top = 1080
End If
End Sub

Private Sub Detail_BeforePrint()
Field3.Text = Field3.Text + " kg"
End Sub

