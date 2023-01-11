Attribute VB_Name = "Module1"

Option Private Module

Public Const GWL_EXSTYLE = -20
Public Const GWL_HINSTANCE = -6
Public Const GWL_HWNDPARENT = -8
Public Const GWL_ID = -12
Public Const GWL_STYLE = -16
Public Const GWL_USERDATA = -21
Public Const GWL_WNDPROC = -4
Public Const DWL_DLGPROC = 4
Public Const DWL_MSGRESULT = 0
Public Const DWL_USER = 8
Public Const NM_CUSTOMDRAW = (-12&)
Public Const WM_NOTIFY  As Long = &H4E&
Public Const CDDS_PREPAINT As Long = &H1&
Public Const CDRF_NOTIFYITEMDRAW As Long = &H20&
Public Const CDDS_ITEM  As Long = &H10000
Public Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Public Const CDRF_NEWFONT As Long = &H2&
Private Const WM_DESTROY = &H2
Public Const WM_NCDESTROY = &H82

Public Handle As Long
Public lpfnOld As Long

#Const DEBUGWINDOWPROC = 0

'#If DEBUGWINDOWPROC Then
    'maintains a WindowProcHook object reference for each subclassed window.
    ' The subclassed window's handle is used as the collection item's key string.
   Public m_colWPHooks As New Collection
'#End If

Public Type NMHDR
   hWndFrom             As Long   ' Window handle of control sending message
   idFrom               As Long        ' Identifier of control sending message
   code                 As Long          ' Specifies the notification code
End Type

' sub struct of the NMCUSTOMDRAW struct

Public Type RECT
   Left                 As Long
   Top                  As Long
   Right                As Long
   Bottom               As Long
End Type

' generic customdraw struct

Public Type NMCUSTOMDRAW
   hdr                  As NMHDR
   dwDrawStage          As Long
   hDC                  As Long
   rc                   As RECT
   dwItemSpec           As Long
   uItemState           As Long
   lItemlParam          As Long
End Type

' listview specific customdraw struct

Public Type NMLVCUSTOMDRAW
   nmcd                 As NMCUSTOMDRAW
   clrText              As Long
   clrTextBk            As Long
   ' if IE >= 4.0 this member of the struct can be used
   'iSubItem            As Integer
End Type

 
Public g_addProcOld     As Long
Public m_BackColorOne As Long
Public m_BackColorTwo As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
    Private Const CLR_INVALID = -1
Public clrTextColor(5000)

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hgdiobj As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Select Case iMsg
      Case WM_NOTIFY
         Dim udtNMHDR         As NMHDR
         CopyMemory udtNMHDR, ByVal lParam, 12&
         With udtNMHDR
            If .code = NM_CUSTOMDRAW Then
               Dim udtNMLVCUSTOMDRAW As NMLVCUSTOMDRAW
               CopyMemory udtNMLVCUSTOMDRAW, ByVal lParam, Len(udtNMLVCUSTOMDRAW)
               With udtNMLVCUSTOMDRAW.nmcd
                  Select Case .dwDrawStage
                     Case CDDS_PREPAINT
                        WindowProc = CDRF_NOTIFYITEMDRAW
                        Exit Function
                     Case CDDS_ITEMPREPAINT
                        If (.dwItemSpec Mod 2) Then
                           If Not (udtNMLVCUSTOMDRAW.clrTextBk = m_BackColorTwo) Then
                                udtNMLVCUSTOMDRAW.clrTextBk = m_BackColorTwo
                           End If
                        Else
                           If Not (udtNMLVCUSTOMDRAW.clrTextBk = m_BackColorOne) Then
                                udtNMLVCUSTOMDRAW.clrTextBk = m_BackColorOne
                           End If
                        End If
                         udtNMLVCUSTOMDRAW.clrText = clrTextColor(.dwItemSpec)
                        CopyMemory ByVal lParam, udtNMLVCUSTOMDRAW, Len(udtNMLVCUSTOMDRAW)
                        WindowProc = CDRF_NEWFONT
                        Exit Function
                  End Select
               End With
            End If
         End With
        
    Case WM_DESTROY
        ' OLDWNDPROC will be gone after UnSubClass is called!
        Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, iMsg, wParam, lParam)
        Call UnSubClass(hWnd)
        Exit Function
    Case WM_NCDESTROY
       Call SetWindowLong(hWnd, GWL_WNDPROC, OLDWNDPROC)
       Exit Function
 
   End Select

    WindowProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, iMsg, wParam, lParam)

End Function

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As Object = Nothing) As Boolean
  
  Dim fSuccess As Boolean
  
  On Error GoTo Out
  
   If GetProp(hWnd, OLDWNDPROC) Then
        SubClass = True
        Exit Function
    End If
  
    #If (DEBUGWINDOWPROC = 0) Then
        lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)
        Handle = hWnd
    #Else
        Set objWPHook = CreateWindowProcHook
        m_colWPHooks.Add objWPHook, CStr(hWnd)
    
        With objWPHook
            Call .SetMainProc(lpfnNew)
            lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
            Call .SetDebugProc(lpfnOld)
        End With

    #End If
  
  If lpfnOld Then
    fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    If (objNotify Is Nothing) = False Then
      fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
    End If
  End If
  
Out:
    If fSuccess Then
        SubClass = True
    Else
        If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)

  If lpfnOld <> 0 Then
    
    'If SetWindowLong(hwnd, GWL_WNDPROC, lpfnOld) Then
      
      Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
      Call RemoveProp(hWnd, OLDWNDPROC)
      Call RemoveProp(hWnd, OBJECTPTR)

#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      On Error Resume Next
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      UnSubClass = True
    
    'End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
   Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If
End Function


