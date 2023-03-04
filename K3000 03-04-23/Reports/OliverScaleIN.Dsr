VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} OliverScaleIN 
   ClientHeight    =   10620
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   25374
   _ExtentY        =   18733
   SectionData     =   "OliverScaleIN.dsx":0000
End
Attribute VB_Name = "OliverScaleIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Detail_BeforePrint()
'Me.Label43.Caption = Format(Field10.Text, "mm-dd-yy")
'Me.Label42.Caption = Format(Field10.Text, "HH:MM am/pm")


'If Val(Me.Field3.Text) > Val(Me.Field9.Text) Then
'gross.Caption = Me.Field3.Text
'tare.Caption = Me.Field9.Text
'Else
'gross.Caption = Me.Field9.Text
'tare.Caption = Me.Field3.Text
'End If
Me.Field3.Text = Me.Field3.Text + " kg"
'Me.Field11.Text = Me.Field11.Text

End Sub


Private Sub GroupHeader1_BeforePrint()
Me.Label39.Caption = Format(Field2.Text, "mm-dd-yy")
Me.Label34.Caption = Format(Field2.Text, "HH:MM am/pm")

End Sub

