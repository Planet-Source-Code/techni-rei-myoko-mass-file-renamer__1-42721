Attribute VB_Name = "SystemIml"
Option Explicit
'
' Brad Martinez,  http://www.mvps.org/ccrp
'
Public Const vbBackslash = "\"
Public Const vbAscDot = 46   ' Asc(".") = 46 (vbKeyDelete)
Public Const vbAllFileSpec = "*.*"

Public Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---


' ===========================================================================
' FindFirstFile

Public Const MAX_PATH = 260

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA   'wfd
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cShortFileName As String * 14
End Type

' WIN32_FIND_DATA.nFileSizeHigh:
' Specifies the high-order DWORD value of the file size, in bytes.
' This value is zero unless the file size is greater than MAXDWORD. The
' size of the file is equal to (nFileSizeHigh * MAXDWORD) + nFileSizeLow.
Public Const MAXDWORD = (2 ^ 32) - 1 ' = 0xFFFFFFFF, not &HFFFFFFFF ( -1)

' FindFirstFile failure rtn value
Public Const INVALID_HANDLE_VALUE = -1

' If the function succeeds, the return value is a search handle
' used in a subsequent call to FindNextFile or FindClose
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
' Rtns True (non zero) on succes, FALSE on failure
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
' Rtns True (non zero) on succes, FALSE on failure
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' ==============================================================
' SHGetFileInfo

'Public Const MAX_PATH = 260

Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

' uFlags (note that the SDK docs state different info):
' Flag that specifies the file information to retrieve.
'   - If uFlags specifies the SHGFI_EXETYPE value, the return value indicates the type
'     of the executable file.
'   - If uFlags includes the SHGFI_SYSICONINDEX value, the return value is the handle
'     to the system image list that contains the specified icon images.
' If uFlags does not include SHGFI_EXETYPE or SHGFI_SYSICONINDEX, the return
' value is nonzero if successful, or zero otherwise.

Public Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum

' ============================================================
' IShellFolder::GetAttributesOf shell attribute flags

Public Const DROPEFFECT_COPY = 1&
Public Const DROPEFFECT_MOVE = 2&
Public Const DROPEFFECT_LINK = 4&

Public Enum SFGAO_Flags

' Capability flags:
  SFGAO_CANCOPY = DROPEFFECT_COPY   ' Objects can be copied
  SFGAO_CANMOVE = DROPEFFECT_MOVE  ' Objects can be moved
  SFGAO_CANLINK = DROPEFFECT_LINK       ' Objects can have shortcuts
  SFGAO_CANRENAME = &H10&                      ' Objects can be renamed
  SFGAO_CANDELETE = &H20&                       ' Objects can be deleted
  SFGAO_HASPROPSHEET = &H40&               ' Objects have property sheets
  SFGAO_DROPTARGET = &H100&                  ' Objects are drop target
  SFGAO_CAPABILITYMASK = &H177&            ' Mask for the capability flags

  ' Display attributes:
  SFGAO_LINK = &H10000                             ' Is a shortcut (link)
  SFGAO_SHARE = &H20000                         ' Is shared
  SFGAO_READONLY = &H40000                  ' Is read-only   *** SHGFI, ISF.GAO for errDeviceNotReady = &H80070015
  SFGAO_GHOSTED = &H80000                    ' Is ghosted icon
  SFGAO_DISPLAYATTRMASK = &HF0000  ' Mask for the display attributes.

  ' Contents attributes (flags):
  SFGAO_FILESYSANCESTOR = &H10000000   ' Is a file system ancestor (local or remoted drive, desktop)
  SFGAO_FOLDER = &H20000000                       ' Is a folder.
  SFGAO_FILESYSTEM = &H40000000                ' Is a file system object (file/folder/root)
  SFGAO_HASSUBFOLDER = &H80000000        ' Expandable in the map pane
  SFGAO_CONTENTSMASK = &H80000000       ' Mask for contents attributes

  ' Miscellaneous attributes:
  SFGAO_VALIDATE = &H1000000           ' invalidate cached information
  SFGAO_REMOVABLE = &H2000000      ' Is removeable media, CD folders do not have this set,
                                                                   ' but child folders of net *drives* do (including mapped
                                                                   ' drives, but not the drive itself nor the net computer)!!
  SFGAO_COMPRESSED = &H4000000   ' Object is compressed (use alt color)

  SFGAO_ALL = &HFFFFFFFF   ' all available attributes, user-defined
End Enum
'

' ==============================================================
' SHGetFileInfo calls

' Obtains and returns the specified information about a file.
'   pszPath  - must be either an absolute path or absolute pidl
'   uFlags    - one or more of SHGFI_ flags
'   sfi           - SHFILEINFO struct passed by calling proc that receives the info
' Returns a value whose meaning depends on the uFlags parameter.
' See each GetFile* proc below that calls this function.

Public Function GetFileInfo(ByVal pszPath As Variant, _
                                            uFlags As Long, _
                                            sfi As SHFILEINFO) As Long
  If (VarType(pszPath) = vbString) Then
    ' Must be an explicit path (not a display name).
    GetFileInfo = SHGetFileInfo(CStr(pszPath), 0, sfi, Len(sfi), uFlags)
  Else   ' assume good pidl
    GetFileInfo = SHGetFileInfo(CLng(pszPath), 0, sfi, Len(sfi), uFlags Or SHGFI_PIDL)
  End If
End Function

' Returns a file's SFGAO_ attributes
'   pszPath  - must be either an absolute path or absolute pidl

Public Function GetFileAttributes(ByVal pszPath As Variant) As Long
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_ATTRIBUTES, sfi) Then
    GetFileAttributes = sfi.dwAttributes
  End If
End Function

' Returns a file's display name (how the file is displayed
' in Explorer, equivalent to calling the GetFileTitle() api),
'   pszPath  - must be either an absolute path or absolute pidl

Public Function GetFileDisplayName(ByVal pszPath As Variant) As String
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_DISPLAYNAME, sfi) Then
    GetFileDisplayName = GetStrFromBufferA(sfi.szDisplayName)
  End If
End Function

' Returns a file's small or large icon index within the system imagelist.
'   pszPath  - must be either an absolute path or absolute pidl
'   uFlags    - either SHGFI_SMALLICON or SHGFI_LARGEICON

Public Function GetFileIconIndex(ByVal pszPath As Variant, uFlags As Long) As Long
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_SYSICONINDEX Or uFlags, sfi) Then
    GetFileIconIndex = sfi.iIcon
  End If
End Function

' Returns the handle of the small or large icon system imagelist.
'   uFlags - either SHGFI_SMALLICON or SHGFI_LARGEICON

Public Function GetSystemImagelist(uFlags As Long) As Long
  Dim sfi As SHFILEINFO
  ' Any valid file system path can be used to retrieve system image list handles.
  GetSystemImagelist = GetFileInfo("C:\", SHGFI_SYSICONINDEX Or uFlags, sfi)
End Function

' ==============================================================
' miscellaneous calls

' Rtns the one-based index of the overlay image shifted left eight bits.

Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  '   INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

Public Function NormalizePath(sPath As String) As String
  If Right$(sPath, 1) <> "\" Then
    NormalizePath = sPath & "\"
  Else
    NormalizePath = sPath
  End If
End Function

Public Function IsFolderAvailable(sFolder As String) As Boolean
  On Error GoTo Out
  IsFolderAvailable = (GetAttr(sFolder) And vbDirectory)
Out:
End Function

