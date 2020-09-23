VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Mass File Renamer"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10905
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView dirtree 
      Height          =   7575
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   13361
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txtindex 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10200
      TabIndex        =   9
      Text            =   "1"
      ToolTipText     =   "This is the number that the ""#""'s will be replaced with"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtpattern 
      Height          =   285
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Click me using a scroll wheel  for more assistance"
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdpattern 
      Caption         =   "&Rename to match pattern"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Tag             =   "0"
      ToolTipText     =   $"frmmain.frx":0E42
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdundo 
      Caption         =   "&Undo"
      Height          =   285
      Left            =   10200
      TabIndex        =   6
      ToolTipText     =   "Change the currently selected file back to what it was"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdrename 
      Caption         =   "&Rename changed file names"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Rename files that you've queued to be renamed"
      Top             =   8040
      Width           =   2655
   End
   Begin VB.FileListBox Filmain 
      Height          =   675
      Hidden          =   -1  'True
      Left            =   1800
      System          =   -1  'True
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "Use this to rename the currently selected file"
      ToolTipText     =   "Currently selected file to rename"
      Top             =   120
      Width           =   7215
   End
   Begin VB.DirListBox Dirmain 
      Height          =   7515
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Use this to select a directory to browse"
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.DriveListBox drvmain 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Use this to select a drive to browse"
      Top             =   120
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   7455
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Use this to select a file to rename"
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   13150
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuapply 
      Caption         =   "Apply"
      Visible         =   0   'False
      Begin VB.Menu mnuapplypattern 
         Caption         =   "&Apply pattern to all files"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuenhance 
         Caption         =   "&Enhance the shell"
      End
   End
   Begin VB.Menu mnurev 
      Caption         =   "Reverse"
      Visible         =   0   'False
      Begin VB.Menu mnureverse 
         Caption         =   "&Reverse direction"
      End
   End
   Begin VB.Menu mnuauto 
      Caption         =   "Auto"
      Visible         =   0   'False
      Begin VB.Menu mnuhelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnusepauto 
         Caption         =   "-"
      End
      Begin VB.Menu mnumacro 
         Caption         =   "&Auto Rename"
         Index           =   0
      End
      Begin VB.Menu mnumacro 
         Caption         =   "&Make unique to a folder"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_hwndTV As Long   ' TreeView.hWnd

Public Sub cmdpattern_Click()
    Dim old(0 To 4) As String
    If txtmain.text <> Empty Then
    
    If mnureverse.Checked = False Then
        cmdpattern.Tag = cmdpattern.Tag + 1
    Else
        cmdpattern.Tag = cmdpattern.Tag - 1
    End If
    If mnureverse.Checked = True Then
        cmdpattern.Tag = cmdpattern.Tag + 1
    End If
    
    old(0) = InStrRev(txtmain.text, ".")
    If Val(old(0)) > 0 Then
        old(0) = Left(txtmain.text, Val(old(0)) - 1)                              'Old Filename
    Else
        old(0) = txtmain.text
    End If
    
    old(1) = InStrRev(txtmain.text, ".")
    If Val(old(1)) > 0 Then
        old(1) = LCase(Right(txtmain.text, Len(txtmain.text) - Val(old(1))))    'Old Extention
    Else
        old(1) = Empty
    End If
    
    old(3) = countchars(txtpattern.text, "#")                                               'Count number of #'s
    
    If old(3) >= Len(cmdpattern.Tag) Then
        old(4) = String(old(3) - Len(cmdpattern.Tag), "0") & cmdpattern.Tag                 'Format index number
    Else
        old(4) = cmdpattern.Tag
    End If
    
    old(2) = txtpattern.text                                                                'Begin replacement
        
    old(2) = Replace(old(2), "&", old(0))                                                           'Filename
    old(2) = Replace(old(2), "%", old(1))                                                           'Extention
    old(2) = Replace(old(2), String(old(3), "#"), old(4))                                           'Index number
    
    If isadir(txtpattern.text) Then
        'filmain.Path is the old dir, txtpattern is the new dir, txtmain.text is the file title
        old(2) = uniquefilename(chkdir(txtpattern, txtmain.text))    ' make the filename unique to the new folder
        old(2) = Right(old(2), Len(old(2)) - InStrRev(old(2), "\")) 'cut it down to the file title
    End If
    
    If old(2) = "@" Then
        'extention must not be altered, cept for ogm
        old(1) = LCase(old(1))
        If old(1) = "ogm" Then old(1) = "avi"
        old(2) = animename(old(0)) & "." & old(1)
        If Right(old(2), 1) = "." Then old(2) = Left(old(2), Len(old(2)) - 1)
    End If
    
    'MsgBox old(0) & vbNewLine & old(1) & vbNewLine & old(2) & vbNewLine & old(3) & vbNewLine & old(4)
    txtmain.text = old(2)
    End If
    txtindex.text = cmdpattern.Tag
End Sub


Private Sub cmdpattern_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Or Button = 4 Then PopupMenu mnuapply
End Sub
Public Function renamefile(source As String, destination As String) As Boolean
On Error Resume Next
Name source As destination
renamefile = True
End Function
Public Function FileExists(Filename As String) As Boolean
On Error Resume Next
    If Dir(Filename, vbNormal + vbHidden + vbSystem + vbDirectory) <> Empty And Filename <> Empty Then FileExists = True Else Filename = False
End Function
Public Sub cmdrename_Click()
On Error GoTo error:
flex.Tag = False
    Dim count As Long
    For count = 1 To flex.Rows - 1
        If getcell(1, count) <> getcell(2, count) Then 'Filename has changed
            'If Replace(Replace(Replace(txtpattern, "#", "?"), "&", "*"), "%", "*") Like getcell(1, count) Then
                'is already like the pattern
            'Else
                If renamefile(chkdir(Filmain.Path, getcell(1, count)), chkdir(Filmain.Path, getcell(2, count))) = True Then
                    setcell 1, count, getcell(2, count)
                    If TextWidth(getcell(2, count)) > flex.ColWidth(1) Then flex.ColWidth(1) = TextWidth(getcell(2, count)) + 240
                End If
            'End If
            DoEvents
        End If
    Next
error:
If Err.Number > 0 Then Call MsgBox("Rename could not complete, a file may be locked, access to the drive may have been lost, or the desired filename is already in use", vbCritical, "Could not rename")
End Sub

Private Sub cmdundo_Click()
    txtmain.text = getcell(1, flex.Row)
End Sub

Private Sub Dirmain_Change()
    On Error Resume Next
    flexclear
    If flex.Tag = True Then If MsgBox("Files were queued to be renamed, would you like to rename them before you leave the directory?", vbQuestion + vbYesNo, "Rename Files?") = vbYes Then cmdrename_Click
    Filmain.Path = Dirmain.Path
End Sub
Public Sub flexclear()
    flex.Rows = 1
    flex.Cols = 3
    flex.ColWidth(0) = 0
    flex.ColWidth(1) = 0
    flex.ColWidth(2) = 0
    setcell 0, 0, "#"
    setcell 1, 0, "File name"
    setcell 2, 0, "New name"
End Sub
Private Sub Dirmain_Click()
    On Error Resume Next
    Dirmain.Path = Dirmain.List(Dirmain.ListIndex)
End Sub

Private Sub dirtree_NodeClick(ByVal Node As ComctlLib.Node)
Dirmain.Path = Node.Tag
End Sub

Public Sub drvmain_Change()
    On Error Resume Next
    flexclear
    If flex.Tag = True Then If MsgBox("Files were queued to be renamed, would you like to rename them before you leave the directory?", vbQuestion + vbYesNo, "Rename Files?") = vbYes Then cmdrename_Click
    Dirmain.Path = drvmain.Drive
    refreshnewgui
End Sub
Public Sub setcell(x As Long, y As Long, text As String)
    If flex.Cols > x And flex.Rows > y Then
        flex.Col = x
        flex.Row = y
        flex.text = text
    End If
End Sub
Public Function getcell(x As Long, y As Long) As String
    If flex.Cols > x And flex.Rows > y Then
        flex.Col = x
        flex.Row = y
        getcell = flex.text
    End If
End Function

Public Sub Filmain_PathChange()
On Error Resume Next
    Dim count As Long, oldpath As String
    flexclear
    Call getcell(0, 0)
    flex.CellBackColor = vbRed
    oldpath = Filmain.Path
    For count = 0 To Filmain.ListCount - 1
        If oldpath <> Filmain.Path Then Exit For 'detect path changes
        
        flex.AddItem count + 1
        setcell 1, flex.Rows - 1, Filmain.List(count): flex.CellAlignment = 0
        setcell 2, flex.Rows - 1, Filmain.List(count): flex.CellAlignment = 0
        If TextWidth(count & "") > flex.ColWidth(0) Then flex.ColWidth(0) = TextWidth(count & "") + 300
        If TextWidth(Filmain.List(count)) > flex.ColWidth(1) Then
            flex.ColWidth(1) = TextWidth(Filmain.List(count)) + 240
            flex.ColWidth(2) = flex.ColWidth(1)
        End If
        DoEvents
        
    Next
    txtmain.text = getcell(1, 1)
    flex.Tag = False
    Call getcell(0, 0)
    flex.CellBackColor = vbGreen
    If flex.Rows > Filmain.ListCount + 1 Then flex.Rows = Filmain.ListCount + 1
End Sub

Public Sub flex_Click()
    txtmain.text = getcell(2, flex.Row)
    txtmain.SelStart = Len(txtmain.text)
    txtmain.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyF5) Then
    MousePointer = vbHourglass
    Call RefreshTreeview(dirtree, dirtree.Nodes(1))
    MousePointer = vbDefault
  End If
End Sub

Private Sub Form_Load()
Dim Filename As String, extention As String, oldfile As String, operand As String, mode As String
If Command <> Empty Then
    Filename = getfromquotes(Command)
    mode = "auto"
    
    If InStr(Filename, ":\") > 2 Then
        operand = LCase(Left(Filename, InStr(Filename, ":\") - 2))
        Filename = Right(Filename, Len(Filename) - InStr(Filename, ":\") + 2)
    End If

    oldfile = Filename
    Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "\"))
    
    mode = IIf(containsword(operand, "box"), "box", mode)
    mode = IIf(containsword(operand, "move"), "move", mode)
    mode = IIf(containsword(operand, "box") And containsword(operand, "auto"), "autobox", mode)
        
    If mode = "autobox" Then Filename = renameinbox(Filename, Left(oldfile, InStrRev(oldfile, "\")) & animename(Filename))
    If mode = "auto" Then Filename = Left(oldfile, InStrRev(oldfile, "\")) & animename(Filename)
    If mode = "box" Then Filename = renameinbox(Filename)
    If mode = "move" Then Filename = movetofolder(Filename)
    
    If containsword(operand, "notify") Then MsgBox oldfile & vbNewLine & IIf(oldfile <> Filename, "was renamed to" & vbNewLine & Filename, " was not renamed. The new and old filenames are identical.")
    If oldfile <> Filename Then renamefile oldfile, Filename
    End
Else
    initnewgui
End If
End Sub

Private Sub Form_Resize()
If Height > 1305 Then
    cmdrename.Top = Height - 780
    Dirmain.Height = Height - 1305
    dirtree.Move Dirmain.Left, Dirmain.Top, Dirmain.Width, Dirmain.Height
    flex.Height = cmdrename.Top + cmdrename.Height - flex.Top
End If

If Width > 3090 Then
    txtindex.Left = flex.Left + flex.Width - txtindex.Width
    cmdundo.Left = flex.Left + flex.Width - cmdundo.Width
    flex.Width = Width - 3090
    txtmain.Width = cmdundo.Left - txtmain.Left - 120
    txtpattern.Width = txtindex.Left - txtpattern.Left - 120
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    uninitnewgui
End Sub

Private Sub mnuapplypattern_Click()
    Dim count As Long
    For count = 1 To flex.Rows - 1
        flex.Row = count
        flex_Click
        cmdpattern_Click
        DoEvents
    Next
    
    If isadir(txtpattern.text) Then
        If MsgBox("Would you like to apply the changes then move all these files to " & txtpattern & "?", vbYesNo, "Move the files") = vbYes Then
            cmdrename_Click
            For count = 1 To flex.Rows - 1
                flex.Row = count
                flex_Click
                copyfile chkdir(Filmain.Path, txtmain.text), chkdir(txtpattern.text, txtmain.text)
            Next
            Filmain.Refresh
            Filmain_PathChange
            If Filmain.ListCount = 0 Then
                If MsgBox("The folder is now empty, would you like to delete it?", vbYesNo, "Delete folder") = vbYes Then
                    RmDir Filmain.Path
                    Filmain.Path = Left(Filmain.Path, InStrRev(Filmain.Path, "\") - 1)
                End If
            End If
        End If
    End If
End Sub
Public Sub copyfile(src As String, dst As String)
    On Error Resume Next 'GoTo err
    FileCopy src, dst
    Kill src
'err:
'    If err.Number > 0 Then MsgBox err.Description, vbCritical, "An error occured"
End Sub
Private Sub mnuenhance_Click()
If MsgBox("This will add some of the renaming features of this program to the shell context menus." & vbNewLine & "(When you right click a file in explorer)" & vbNewLine & "It will ask you which ones you wish to add" & vbNewLine & "Would you like to continue?", vbYesNo, "Shell Enhancements") = vbYes Then
    If MsgBox("Adds the ability to automatically rename a file, like the '@' pattern does", vbYesNo, "Shell Enhancement: Auto Rename") = vbYes Then SaveString HKEY_CLASSES_ROOT, "*\Shell\Auto Rename\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" %1"
    If MsgBox("An input box is shown with the filename for you to change. I find this easier than the listview's rename ability", vbYesNo, "Shell Enhancement: Rename in box") = vbYes Then SaveString HKEY_CLASSES_ROOT, "*\Shell\Rename in box\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" box %1"
    If MsgBox("Move the file to another directory with a unique filename", vbYesNo, "Shell Enhancement: Move with unique filename") = vbYes Then SaveString HKEY_CLASSES_ROOT, "*\Shell\Move\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" move %1"
    If MsgBox("Similar to 'Rename in box', but the default option is auto renamed like the '@' pattern", vbYesNo, "Shell Enhancement: Auto Rename in box") = vbYes Then SaveString HKEY_CLASSES_ROOT, "*\Shell\Auto Rename in box\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" autobox %1"
End If
End Sub

Private Sub mnuhelp_Click()
MsgBox Replace("Use '#' to represent the index number.%nThe number of '#''s will represent the length of the index number.%nUse '%' to represent to old extention%n'&' to represent to old file name.%nIf you put only '@' in this, it will auto rename%nIf a directory is placed here, the folder/directory will be made unique to that folder%n%nAlso, right click 'Rename to match Pattern' to apply it to all the files in the current directory", "%n", vbNewLine), vbInformation, "Instructions"
End Sub

Private Sub mnumacro_Click(index As Integer)
If index = 0 Then txtpattern = "@"
If index = 1 Then txtpattern = BrowseForFolder(Me.hWnd, "Please select the folder")
End Sub

Private Sub mnureverse_Click()
mnureverse.Checked = Not mnureverse.Checked
End Sub

Private Sub txtindex_Change()
If txtindex <> Empty Then
cmdpattern.Enabled = True
cmdrename.Enabled = True
    If IsNumeric(txtindex) = False Or txtindex < 1 Then txtindex = 1
    cmdpattern.Tag = Val(txtindex.text - 1)
Else
cmdpattern.Enabled = False
cmdrename.Enabled = False
End If
End Sub

Private Sub txtindex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then PopupMenu mnurev
End Sub

Private Sub txtmain_Change()
    If flex.Row > 0 Then setcell 2, flex.Row, Trim(txtmain.text)
    If TextWidth(txtmain.text) > flex.ColWidth(2) Then flex.ColWidth(2) = TextWidth(txtmain.text) + 240
    If txtmain.text <> getcell(1, flex.Row) Then flex.Tag = True
End Sub

Private Sub txtpattern_Change()
    cmdpattern.Enabled = True
    cmdpattern.Tag = 0
    txtindex.text = 1
End Sub
Public Sub initnewgui()
  KeyPreview = True   ' for TreeView refresh
  ' Initialize the TreeView...
  With dirtree
    .HideSelection = False
    .Indentation = 19 * Screen.TwipsPerPixelX  ' default common control treeview indentation.
    .LabelEdit = tvwManual
    m_hwndTV = .hWnd
  End With
  Set treeloc = dirtree
  Call TreeView_SetImageList(m_hwndTV, GetSystemImagelist(SHGFI_SMALLICON), TVSIL_NORMAL)
  Call SubClass(m_hwndTV, AddressOf TVWndProc)
  Call SubClass(hWnd, AddressOf FrmWndProc)
  flex.Tag = False
  Call drvmain_Change
  Dirmain.Path = dirtree.Nodes(1).Tag
End Sub
Public Sub uninitnewgui()
  Call UnSubClass(m_hwndTV)
  Call TreeView_SetImageList(m_hwndTV, 0, TVSIL_NORMAL)
  Call RemoveRootFolder(dirtree)
End Sub
Public Sub refreshnewgui()
  Dim sdrive As String
  Static sPrevDrive As String
  sdrive = UCase(Left$(drvmain.Drive, 2) & "\")
  If (sdrive <> sPrevDrive) Then
    If IsFolderAvailable(sdrive) Then
      sPrevDrive = sdrive
      Call InsertRootFolder(dirtree, sdrive)
    Else
      Call MsgBox(sdrive & " currently is unavailable.", vbExclamation, "Unable to load drive contents")
      drvmain.Drive = sPrevDrive
    End If   ' IsFolderAvailable
  End If   ' (sDrive <> sPrevDrive)
End Sub

Private Sub txtpattern_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 4 Then PopupMenu mnuauto
End Sub
