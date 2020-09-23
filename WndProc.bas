Attribute VB_Name = "mWndProc"
Option Explicit
Public treeloc As Object
' Brad Martinez,  http://www.mvps.org/ccrp
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
'
' ==============================================
' A general purpose subclassing module w/ debugging code
' ==============================================

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

#If DEBUGWINDOWPROC Then
  ' maintains a WindowProcHook reference for each subclassed window.
  ' the window's handle is the collection item's key string.
  Private m_colWPHooks As New Collection
#End If
'

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As Object = Nothing) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out

  If GetProp(hWnd, OLDWNDPROC) Then
    SubClass = True
    Exit Function
  End If
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)

#Else
    Dim objWPHook As WindowProcHook
    
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
    MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
                  "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)
  If lpfnOld Then
    
    If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
      Call RemoveProp(hWnd, OLDWNDPROC)
      Call RemoveProp(hWnd, OBJECTPTR)

#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      UnSubClass = True
    
    End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

' Processes Form1 window messages

Public Function FrmWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg
      
    ' ============================================================
    ' Process treeview notification messages
    
    Case WM_NOTIFY
      Dim nmh As NMHDR
      Dim nmtv As NMTREEVIEW
      Dim nod As Node
      
      ' If a treeview item is about to expand...
      MoveMemory nmh, ByVal lParam, Len(nmh)
      If (nmh.code = TVN_ITEMEXPANDING) Then
        
        ' Fill the NMTREEVIEW struct. If the struct's itemOld.pszText and
        ' itemNew.pszText members were defined as strings, and we didn't
        ' pre-allocate both, we'd GPF...
        MoveMemory nmtv, ByVal lParam, Len(nmtv)
        
        ' If the expanding folder has no children (the VB
        ' TreeView does not set TVIS_EXPANDEDONCE)
'        If ((nmtv.itemNew.state And TVIS_EXPANDEDONCE) = False) Then
        If (TreeView_GetChild(nmh.hwndFrom, nmtv.itemNew.hItem) = 0) Then
          
          ' We'll cheat and go undoc to get a Node reference...
          ' (not recommended for production code...)
          Set nod = GetNodeFromlParam(nmtv.itemNew.lParam)
          If (nod Is Nothing) = False Then
            Call InsertSubfolders(treeloc, nmtv.itemNew.hItem, nod, nod.Tag)
          End If
        End If   ' TreeView_GetChild
      End If   ' (nmh.code = TVN_ITEMEXPANDING)
      
    ' ============================================================
    ' Unsubclass the window if we forget to do it... (and yes, we did)
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
  
  End Select
  
  FrmWndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)

End Function

' Processes treeview window messages.

Public Function TVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg

    ' ============================================================
    ' Prevent the TreeView from removing our system imagelist assignment, which
    ' it wil do when it sees no VB ImageList associated with it.
    ' (the TreeView can't be subclassed when we're assigning imagelists...)
    
    Case TVM_SETIMAGELIST
      Exit Function
          
    ' ============================================================
    ' Unsubclass the window if we forget to do it...
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
  
  End Select
  
  TVWndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)

End Function
