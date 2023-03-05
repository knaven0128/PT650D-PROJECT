Attribute VB_Name = "GenCode"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'active connection and command
Public con As New ADODB.Connection
Public ocn As New ADODB.Connection
Public cmd As New ADODB.Command
'recordsets
Public rsexport As New ADODB.Recordset
Public rsuser As New ADODB.Recordset
Public rsproduct As New ADODB.Recordset
Public rstransac As New ADODB.Recordset
Public rstruck As New ADODB.Recordset
Public rstrucking As New ADODB.Recordset
Public rscustomer As New ADODB.Recordset
Public rslog As New ADODB.Recordset
Public rs As New ADODB.Recordset
Public rsdel As New ADODB.Recordset
Public rsdb As New ADODB.Recordset
Public rsconnect As New ADODB.Recordset
Public rslocate As New ADODB.Recordset
Public rscomm As New ADODB.Recordset
Public rscompany As New ADODB.Recordset
Public rsum As New ADODB.Recordset
Public rsscaleuser As New ADODB.Recordset
Public rsserial As New ADODB.Recordset
Public rsdata As New ADODB.Recordset
Public rscount As New ADODB.Recordset
Public rsexcel As New ADODB.Recordset
Public rslogs As New ADODB.Recordset
Public rsdest As New ADODB.Recordset
Public dbloc As String

Public loc As String
Public settings As String, ports As String, RThreshold As String, currentuser As String, currentposition As String, logctrs As String, lognum As String, currentname As String
Public cnt As Integer, entrycus As Byte, consec_num As String, count_num As String
Public sDateString As String, companyname As String
Public serial As String
Public register As String, perm As Byte, Cnumber As String
Public oExcel As Object
Public oBook As Object
Public oSheet As Object
Public DataArray(1 To 999, 1 To 9999) As Variant
Public R As Integer
Public NumberOfRows As Integer
Public listnum As Byte
Public commlen As String
Public commstr As String
Public servno As String, userr As String, passs As String, comset As String, connectt As String, logid As String, commsymbol As String, commpositive As String, commnegative As String
Public comnum As Integer
'Public Function openn() As Boolean
'    If con.State = adStateOpen Then con.Close
'  con.CursorLocation = adUseClient
''con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbloc & ";Persist Security Info=False;JET OLEDB:Database Password=Babyrr0403"
'con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\weighingdb.mdb;Persist Security Info=False;JET OLEDB:Database Password=Babyrr0403"
'con.Open
'    If con.State = adStateOpen Then
'        openn = True
'    End If
'
'End Function


Public Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long
 
    cmd = "cmd.exe /cd " & cmd
    result = Shell(cmd, vbHide)
 
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

Public Function openndb() As Boolean
    If ocn.State = adStateOpen Then ocn.Close
  ocn.CursorLocation = adUseClient
'con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbloc & ";Persist Security Info=False;JET OLEDB:Database Password=Babyrr0403"
ocn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\weighingdb.mdb;Persist Security Info=False;JET OLEDB:Database Password=Babyrr0403"
ocn.Open
    If ocn.State = adStateOpen Then
        openndb = True
    End If

End Function
Public Function openn() As Boolean

    If con.State = adStateOpen Then con.Close
    
    con.CursorLocation = adUseClient
    con.Mode = adModeShareDenyNone
    con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=" & servno & ";" _
        & "DATABASE=weighingscaledb;" _
        & "UID=" & userr & ";" _
        & "PWD=" & passs & ";" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841    'SET ALL PARAMETERS
    
    con.Open
    If con.State = adStateOpen Then
        openn = True
    End If
        
End Function


Public Function comm() As Boolean
If frmWT.MSComm1.settings = " & connectt & " Then
ElseIf frmWT.MSComm1.CommPort = " & comportt & " Then
End If

End Function
        
Public Sub Emptyctl(X As Form, tg As String)
    Dim c As Control
    
    For Each c In X.Controls
        If (TypeOf c Is TextBox) And ((c.Tag = tg) Or (c.Tag = "txtcom")) Then
            c.Text = ""
        ElseIf (TypeOf c Is ComboBox) And (c.Tag = tg) Then
            c.Refresh
        End If
    Next
End Sub
Public Sub popgrid(dG As DataGrid, s As ADODB.Recordset)
    Set dG.DataSource = Nothing
    Set dG.DataSource = s.DataSource
End Sub
Public Sub Charctl(X As Form, tg As String, KeyAscii As Integer)
Dim c As Control
For Each c In X.Controls
If (TypeOf c Is TextBox) And ((c.Tag = tg) Or (c.Tag = "txtcom")) Then
Select Case KeyAscii
Case vbKeyReturn And 65 To 90, 48 To 57, 8 ' A-Z, 0-9 and backspace
'Let these key codes pass through
Case vbkeyesc And 97 To 122, 8 'a-z and backspace
'Let these key codes pass through
Case Else
'All others get trapped
MsgBox "Supply Correct Data.", vbExclamation, "Error:"
KeyAscii = 0 ' set ascii 0 to trap others input
End Select
End If
 Next
End Sub
Public Function generate(A As String, numm As Integer) As String
    Dim lenn As Integer, leftt As Integer, B As Integer
        
    lenn = Len(A)
    
    leftt = numm - lenn
    
    For B = 1 To leftt
        A = "0" & A
    Next B
    
    
    If numm = 12 Then 'OR 15 soldid
        generate = "OR-" & A
    Else
        generate = A
    End If
        
End Function
Public Function generate1(A As String, numm As Integer) As String
    Dim lenn As Integer, leftt As Integer, B As Integer
        
    lenn = Len(A)
    
    leftt = numm - lenn
    
    For B = 1 To leftt
        A = "0" & A
    Next B
    
    
    If numm = 12 Then 'OR 15 soldid
        generate1 = "OR-" & A
    Else
        generate1 = A
    End If
        
End Function
Public Sub cmmd(c As String)
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = con
        .CommandText = c
    End With
    Call cmd.Execute
End Sub
Public Sub cmmd1(c As String)
    With cmd
        .CommandType = adCmdText
        .ActiveConnection = ocn
        .CommandText = c
    End With
    Call cmd.Execute
End Sub
Public Sub numonly(KeyAscii As Integer)
    Select Case VBA.Chr(KeyAscii)
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ".", "-"
            KeyAscii = Asc(VBA.Chr(KeyAscii))
        Case VBA.Chr(vbKeyBack)
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub


Public Sub DrawProgress(picProgress As PictureBox, ByVal Value As Long, _
    Optional lngMin As Long = 0, Optional lngMax As Long = 100, _
    Optional lngBackColor As Long = vbWhite, Optional lngForeColor As Long = vbBlue)
    
Dim strPercent As String
Dim intX As Integer
Dim intY As Integer
Dim intWidth As Integer
Dim intHeight As Integer
Dim sngPercent As Single
    
    sngPercent = Int((CSng(Value) / (CSng(lngMax) - CSng(lngMin))) * 100)
    strPercent = CStr(sngPercent) & "%"

    intWidth = picProgress.TextWidth(strPercent)
    intHeight = picProgress.TextHeight(strPercent)
    intX = (picProgress.ScaleWidth - intWidth) / 2
    intY = (picProgress.ScaleHeight - intHeight) / 2
    
    With picProgress
        .AutoRedraw = True
        .FillStyle = vbFSSolid
        .BackColor = lngBackColor
        .ForeColor = lngForeColor
        .DrawMode = vbCopyPen
        
        .CurrentX = intX
        .CurrentY = intY
        picProgress.Print strPercent
        .DrawMode = vbNotXorPen
    End With
    
    If sngPercent > 0 Then
        picProgress.Line (0, 0)-(picProgress.width * sngPercent / 100, picProgress.height), lngForeColor, BF
    Else
        picProgress.Line (0, 0)-(picProgress.width, picProgress.height), lngBackColor, BF
    End If
    picProgress.Refresh
End Sub
