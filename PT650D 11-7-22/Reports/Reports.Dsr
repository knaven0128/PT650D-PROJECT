VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Reports 
   Caption         =   "Reports"
   ClientHeight    =   12375
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "Reports.dsx":0000
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_BeforePrint()
Me.Field4.Text = Me.Field4.Text + " kg"
Me.Field5.Text = Me.Field5.Text + " kg"
Me.Field6.Text = Me.Field6.Text + " kg"
Me.Field14.Text = Me.Field14.Text + " kg"
If Val(Field4.Text) > Val(Field5.Text) Then
Me.Field4.Left = 5295
Me.Field5.Left = 6465
Else
Me.Field5.Left = 5295
Me.Field4.Left = 6465
End If

End Sub

