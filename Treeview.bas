Attribute VB_Name = "mTreeview"
Option Explicit

' ==============================================================
' treeview definitions

Public Type TVITEM   ' was TV_ITEM
  mask As Long
  hItem As Long
  state As Long
  stateMask As Long
  pszText As Long    ' if a string, must be pre-allocated when struct is filled!
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

' TVITEM mask flags.
Public Const TVIF_TEXT = &H1
Public Const TVIF_IMAGE = &H2
Public Const TVIF_STATE = &H8
Public Const TVIF_SELECTEDIMAGE = &H20
Public Const TVIF_CHILDREN = &H40

' TVITEM state and statemask.
Public Const TVIS_OVERLAYMASK = &HF00

' treeview window messages
Public Const TV_FIRST = &H1100
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Public Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Public Const TVM_SETITEM = (TV_FIRST + 13)

' TVM_SETIMAGELIST wParam
Public Const TVSIL_NORMAL = 0

' TVM_GETNEXTITEM wParam
Public Const TVGN_ROOT = &H0
Public Const TVGN_NEXT = &H1
Public Const TVGN_CHILD = &H4

' treeview notification messages
Public Const TVN_FIRST = -400&    ' (0U-400U) or &HFFFFFE70
Public Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)   ' lParam = lpNMTREEVIEW

Public Type NMTREEVIEW   ' was NM_TREEVIEW
  hdr As NMHDR
  action As Long
  itemOld As TVITEM
  itemNew As TVITEM
  ptDrag As POINTAPI
End Type
'

' ==============================================================
' treeview calls

' Inserts a new root folder into the treeview control.
'
'   objTV    - treeview reference
'   sFolder  - root item's fully qualified path
'
' If successful, returns the root folder's treeview item handle, returns 0 otherwise.
'
' Called only from Form1.Drive1_Change

Public Function InsertRootFolder(objTV As TreeView, sFolder As String) As Long
  Dim hitemRoot As Long
    
  Call RemoveRootFolder(objTV)
    
  ' Insert the root
  hitemRoot = InsertFolder(objTV, Nothing, 0, 0, sFolder)
  If hitemRoot Then

    ' Expand the root, invoking a TVN_ITEMEXPANDING,
    ' which calls InsertSubfolders() for the root's subfolders.
    objTV.Nodes(1).Expanded = True
    objTV.Nodes(1).Selected = True
    
    InsertRootFolder = hitemRoot
  End If

End Function

' Removes the root folder and all of its subfolders from the specified treeview

'   objTV   - treeview reference

' called from InsertRootFolder above and Form1.Form_Unload

Public Sub RemoveRootFolder(objTV As TreeView)
    
  If objTV.Nodes.count Then
    objTV.Nodes(1).Root.Expanded = False
    Call objTV.Nodes.Remove(objTV.Nodes(1).Root.Index)
  End If
  
End Sub

' Inserts the specified folder under the specified parent folder
'
'   hwndTV           - treeview's hWnd
'   nodParent        - parent folder's Node reference
'   hitemParent      - parent folder's treeview item handle, is 0 for root folder
'   hitemPrevChild - parent folder's previous child's treeview item handle, is 0 for parent's first child
'   sFolder             - fully qualified path of the folder being inserted.

' If successful, returns the folder's treeview item handle, returns 0 otherwise.

' Called from InsertRootFolder above, and InsertSubfolders below.

Public Function InsertFolder(objTV As TreeView, _
                                            nodParent As Node, _
                                            hitemParent As Long, _
                                            hitemPrevChild As Long, _
                                            sFolder As String) As Long
  Dim ulAttrs As Long
  Dim tvi As TVITEM
  Dim nod As Node
  
  ' Get the folder's attributes
  ulAttrs = GetFileAttributes(sFolder)
  
  ' Though FindFirstFile will enumerate virtual folders (i.e. History subfolders)
  ' SHGetFileInfo can only evaluate virtual folders by their PIDLs. Since we're
  ' not doing PIDLs, we'll bail in this situation....
  If (ulAttrs = 0) Then Exit Function
  
  ' Indicate what TVITEM members will contain data
  tvi.mask = TVIF_CHILDREN Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
  
  ' ====================================================
  ' By explicitly setting the treeview item attributes that the VB TreeView
  ' normally does callbacks for, we increase the performance of the
  ' TreeView dramatically, since a TVN_GETDISPINFO is sent by the real
  ' treeview any time, any item, needs to present any of these attributes.
  ' One problem though, this information is not available in the Node...
  ' ...an insignificant side effect since this info can be obtained by APIs...
  
  ' If the folder has subfolders, explicitly give the item a button, overriding
  ' the I_CHILDRENCALLBACK value the VB TreeView uses.
  tvi.cChildren = Abs(CBool(ulAttrs And SFGAO_HASSUBFOLDER))
  
  ' Explicitly set the folder's normal and selected icon indices, overriding
  ' the I_IMAGECALLBACK value the VB TreeView uses.
  tvi.iImage = GetFileIconIndex(sFolder, SHGFI_SMALLICON)
  tvi.iSelectedImage = GetFileIconIndex(sFolder, SHGFI_SMALLICON Or SHGFI_OPENICON)

  ' If the folder is shared, give it the share overlay icon.
  If (ulAttrs And SFGAO_SHARE) Then
    tvi.mask = tvi.mask Or TVIF_STATE
    tvi.state = TVIS_OVERLAYMASK
    ' share overlay is the 1st system imagelist overlay image,
    ' shortcut is 2nd, gray arrow is 3rd, no 4th image
    tvi.stateMask = INDEXTOOVERLAYMASK(1)
  End If
  
#If (USEDISPINFO = 0) Then
  ' Explicitly set the folder's text, overriding the LPSTR_TEXTCALLBACK
  ' value the VB TreeView uses.
  tvi.mask = tvi.mask Or TVIF_TEXT
  tvi.pszText = StrPtr(String$(MAX_PATH, 0))
  Call lstrcpyA(ByVal tvi.pszText, ByVal GetFileDisplayName(sFolder))
  tvi.cchTextMax = MAX_PATH
  
  ' Add the Node to the TreeView, without a button, icons, or text (we did everything
  ' above first so that there's the least amount of code between inserting the Node
  ' and setting its attributes below). This must happen, since the TreeView needs a
  ' place-holder for all items in the tree.
  If (nodParent Is Nothing) Then
    Set nod = objTV.Nodes.Add
  Else
    Set nod = objTV.Nodes.Add(nodParent, tvwChild)
  End If
  
#Else
  ' Set the Node's Text. The real treeview will send a TVN_GETDISPINFO each
  ' time it needs an item's text. The VB TreeView's code will then read and specify
  ' the Node.Text in response.
  If (nodParent Is Nothing) Then
    Set nod = objTV.Nodes.Add(, , , GetFileDisplayName(sFolder))
  Else
    Set nod = objTV.Nodes.Add(nodParent, tvwChild, , GetFileDisplayName(sFolder))
  End If
  
#End If   ' (USEDISPINFO = 0)

  ' Since we may be setting the item's Text, making Node.Text
  ' empty, and since we're using drive's displaynames instead of
  ' their paths for Node.Text, we'll store the item's path in the
  ' Node's Tag and not even deal with Node.FullPath...
  nod.Tag = sFolder
  
  ' ====================================================
  
  ' Get the new Node's hItem
  If (hitemParent = 0) Then
    tvi.hItem = TreeView_GetRoot(objTV.hWnd)
  ElseIf (hitemPrevChild = 0) Then
    tvi.hItem = TreeView_GetChild(objTV.hWnd, hitemParent)
  Else
    tvi.hItem = TreeView_GetNextSibling(objTV.hWnd, hitemPrevChild)
  End If
  
  ' And set the item's button, icons (and text), done deal...
  Call TreeView_SetItem(objTV.hWnd, tvi)
  
  ' Return the folder's hItem
  InsertFolder = tvi.hItem

End Function

' Inserts subfolders under the specified parent folder.

'   hwndTV        - treeview hwndOwner
'   hitemParent   - parent folder's treeview item handle
'   nodParent     - parent folder's Node reference
'   sParentPath  - parent folder's fully qualified path

' Called only from FrmWndProc/TVN_ITEMEXPANDING

Public Sub InsertSubfolders(objTV As TreeView, _
                                            hitemParent As Long, _
                                            nodParent As Node, _
                                            ByVal sParentPath As String)
  Dim hwndOwner As Long
  Dim wfd As WIN32_FIND_DATA
  Dim hFind As Long
  Dim sSubfolder As String
  Dim hitemChild As Long
  Dim tvi As TVITEM
  
  hwndOwner = GetParent(objTV.hWnd)
  Screen.MousePointer = vbHourglass
  
  ' make sure the parent path has a trailing backslash...
  sParentPath = NormalizePath(sParentPath)
  
  ' Make sure Nodes aren't sorted each time one is inserted...
  nodParent.Sorted = False
  
  hFind = FindFirstFile(sParentPath & vbAllFileSpec, wfd)
  If (hFind <> INVALID_HANDLE_VALUE) Then
  
    Do
      
      If (wfd.dwFileAttributes And vbDirectory) Then
        ' If not a  "." or ".." DOS subdir...
        If (Asc(wfd.cFileName) <> vbAscDot) Then
          ' Append the subfolder's filename to the parent folder's
          ' path and insert it under the parent folder.
          hitemChild = InsertFolder(objTV, nodParent, hitemParent, hitemChild, _
                                                 sParentPath & Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1))
        End If
      End If   ' (wfd.dwFileAttributes And vbDirectory)
      
    Loop While FindNextFile(hFind, wfd)
  
    Call FindClose(hFind)
  End If  ' (hFind <> INVALID_HANDLE_VALUE)
  
  ' If for some reason we didn't load any subfolders under the parent
  ' folder, remove the parent folder's button...
  If (hitemChild = 0) Then
    tvi.hItem = hitemParent
    tvi.mask = TVIF_CHILDREN
    tvi.cChildren = 0
    Call TreeView_SetItem(objTV.hWnd, tvi)
  Else
    ' Sort the parent folder's subfolders alphabetically.
    nodParent.Sorted = True
  End If
  
  Screen.MousePointer = vbDefault

End Sub

' Refreshes the treeview by removing all subfolders under all collapsed paent folders

'   objTV       - treeview's hWnd
'   nodSibling  - Node reference of the first sibling under any given parent Node
'                        On first call, pass TreeView.Nodes(1) for this param.

' called from Form_KeyDown

Public Sub RefreshTreeview(objTV As TreeView, nodSibling As Node)
  Dim nodChild  As Node
  
  Do While (nodSibling Is Nothing) = False

    ' remove all children of collapsed sibling Nodes
    If nodSibling.Expanded Then
      Call RefreshTreeview(objTV, nodSibling.Child)
    Else
      ' nodSibling.Children calls TVM_GETNEXTITEM/TVGN_NEXTs for the whole
      ' sibling hierarchy, this method sends the least amount of TVM_GETNEXTITEMs...
      ' And be sure that the parent Node Sorted = False before re-inserting children...
      Set nodChild = nodSibling.Child
      Do While (nodChild Is Nothing) = False
        objTV.Nodes.Remove nodChild.Index
        Set nodChild = nodSibling.Child
      Loop
    End If

    Set nodSibling = nodSibling.Next
  Loop

End Sub

Public Function GetNodeFromlParam(lParam As Long) As Node
  Dim pNode As Long
  Dim nod As Node
  
  MoveMemory pNode, ByVal lParam + 8, 4
  If pNode Then
    MoveMemory nod, pNode, 4
    Set GetNodeFromlParam = nod
    MoveMemory nod, 0&, 4
  End If
  
End Function

' ==============================================================
' treeview macros

' Sets the normal or state image list for a tree-view control and redraws the control using the new images.
' Returns the handle to the previous image list, if any, or 0 otherwise.

Public Function TreeView_SetImageList(hWnd As Long, hIml As Long, iImage As Long) As Long
  TreeView_SetImageList = SendMessage(hWnd, TVM_SETIMAGELIST, iImage, ByVal hIml)
End Function

' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Public Function TreeView_SetItem(hWnd As Long, pitem As TVITEM) As Boolean
  TreeView_SetItem = SendMessage(hWnd, TVM_SETITEM, 0, pitem)
End Function

' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetRoot(hWnd As Long) As Long
  TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function

' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
  TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
  TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)
End Function

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
  TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function
