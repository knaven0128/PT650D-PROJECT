VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SetListViewAlternateColor(ByRef objlistview As Object)

    On Error GoTo Err_SetLstViewColor

    Dim sRowHeight As Single
    Dim sScaleWidth As Single
    Dim icount As Integer
    Dim lvHWND As Long
    Dim lItem As Long

    On Error GoTo Err_SetLstViewColor

    If TypeName(objlistview) = "ListView" Then

        m_BackColorOne = &HFFFFFF
        m_BackColorTwo = &HC0FFFF

        'Call SubClass(GetParent(objlistview.hWnd), AddressOf WindowProc)
        'Handle = GetParent(objlistview.hWnd)
        
        For lItem = 1 To objlistview.ListItems.Count - 1
            SetListViewItemForeColor objlistview, lItem, TranslateColor(objlistview.ListItems(lItem).ForeColor)
        Next lItem

        
         Call SubClass(GetParent(objlistview.hWnd), AddressOf WindowProc)
       ' Handle = GetParent(objlistview.hWnd)

    ElseIf TypeName(objlistview) = "fpList" Then
            objlistview.ListApplyTo = 6 'ListApplyToEvenRows
            objlistview.BackColor = RowAlternateColor1
            objlistview.ListApplyTo = 5 'ListApplyToOddRows
            objlistview.BackColor = RowAlternateColor2
    End If

    Exit Sub

Err_SetLstViewColor:
 ' err.Raise err.Number, "SetLstViewColor", err.Description
End Sub

Private Sub SetListViewItemForeColor(lvw As ListView, lItem As Long, strColor As Long)
    
   On Error GoTo Err_SetListViewItemForeColor
    With lvw
        If .ListItems.Count = 0 Then Exit Sub
        'ReDim Preserve clrTextColor(lItem)
        clrTextColor(lItem - 1) = strColor
        .Refresh
    End With

Exit Sub

Err_SetListViewItemForeColor:
    Err.Raise Err.number, "Err_SetListViewItemForeColor", Err.Description
End Sub

