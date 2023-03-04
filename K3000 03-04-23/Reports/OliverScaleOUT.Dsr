VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} OliverScaleOUT 
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19288
   SectionData     =   "OliverScaleOUT.dsx":0000
End
Attribute VB_Name = "OliverScaleOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Detail_BeforePrint()
Me.Label43.Caption = Format(Field10.Text, "mm-dd-yy")
Me.Label42.Caption = Format(Field10.Text, "HH:MM am/pm")


If Val(Me.Field3.Text) > Val(Me.Field9.Text) Then
gross.Caption = Me.Field3.Text
tare.Caption = Me.Field9.Text
Else
gross.Caption = Me.Field9.Text
tare.Caption = Me.Field3.Text
End If
Me.Field3.Text = Me.Field3.Text + " kg"
Me.Field11.Text = Me.Field11.Text

End Sub

Private Sub Detail_Format()
If Me.Field19.Text > 0 Then
Me.Field19.Text = "PHP " + Format(Me.Field19.Text, "###,###,##0.00")
Else
Me.Field19.Text = ""
End If
If Me.Field20.Text > 0 Then
Me.Field20.Text = "PHP " + Format(Me.Field20.Text, "###,###,##0.00")
Else
Me.Field20.Text = ""
End If
If Me.Field7.Text > 0 Then
Me.Field7.Text = Format(Me.Field7.Text, "###,###,##0.00")
Else
Me.Field7.Text = ""
End If
If Me.Field14.Text > 0 Then
Me.Field14.Text = Format(Me.Field14.Text, "###,###,##0.00")
Else
Me.Field14.Text = ""
End If

If Me.Field13.Text <= 0 Then
Me.Field13.Text = ""
End If
End Sub

Private Sub GroupHeader1_BeforePrint()
Me.Label39.Caption = Format(Field2.Text, "mm-dd-yy")
Me.Label34.Caption = Format(Field2.Text, "HH:MM am/pm")
If Field12.Text = "NA" Or Field12 = "None" Then
Field12.Visible = False
Else
Field12.Visible = True

End If



End Sub



