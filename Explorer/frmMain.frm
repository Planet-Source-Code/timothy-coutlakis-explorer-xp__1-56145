VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Explorer"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   Begin VB.PictureBox picBuffer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2040
      ScaleHeight     =   360
      ScaleWidth      =   465
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSComctlLib.ImageList ImgLarge 
      Left            =   720
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgSmall 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0562
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameDummyDrives 
      Caption         =   "Drives"
      Height          =   2895
      Left            =   3840
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   4455
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   2280
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox DummyDir 
         Height          =   990
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   990
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
      End
      Begin VB.DirListBox CheckForChildDir 
         Height          =   990
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   767
      ButtonWidth     =   820
      ButtonHeight    =   714
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Up"
            ImageKey        =   "Up"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2880
      Left            =   3120
      ScaleHeight     =   1254.076
      ScaleMode       =   0  'User
      ScaleWidth      =   624
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.CommandButton cmdClosePreview 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3735
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0E72
   End
   Begin MSComctlLib.ListView lvDir 
      DragIcon        =   "frmMain.frx":0EFD
      Height          =   4935
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8705
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImgLarge"
      SmallIcons      =   "ImgSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.TreeView tvDir 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8705
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   3
      ImageList       =   "ImgSmall"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   1320
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":133F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197B
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FB7
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   2865
      Left            =   2880
      MousePointer    =   9  'Size W E
      Top             =   360
      Width           =   150
   End
   Begin VB.Image pPicture 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line LinePos 
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Visible         =   0   'False
         Begin VB.Menu mnuNewFolder 
            Caption         =   "Folder..."
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview..."
      End
      Begin VB.Menu mnuPreviewAs 
         Caption         =   "Preview As"
         Begin VB.Menu mnuPreviewText 
            Caption         =   "Text Document"
         End
         Begin VB.Menu mnuPreviewPicture 
            Caption         =   "Picture"
         End
         Begin VB.Menu mnuPreviewVideo 
            Caption         =   "Video Or Sound"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Type PointApi
  x As Long
  y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMRECT = LVM_FIRST + 14
Private Const LVIR_BOUNDS = 0

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointApi) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim i As Integer
Dim MeOpen As Boolean
Dim MouseButtonPressed As Integer
Dim lvButton As Integer
Dim lvMouseDown As Boolean
Dim lvMousePos As PointApi
Dim PopUpMnu As Boolean

Dim thisFile As File
Dim thisFolder As Folder

Dim fSys As FileSystemObject

Dim mbMoving As Boolean
Const sglSplitLimit = 50

Enum PreviewType
  [pvText]
  [pvpicture]
  [pvVideo]
End Enum

Dim key_Shift As Boolean

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "Comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim FileInfo As typSHFILEINFO

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer, iKey As String) As Long
  Dim SmallIcon As Long
  Dim NewImage As ListImage
  Dim IconIndex As Integer
  
  On Error GoTo Load_New_Icon
  
  If iKey <> "Application" And iKey <> "Shortcut" Then
    ExtractIcon = AddtoImageList.ListImages(iKey).Index
    Exit Function
  End If
  
Load_New_Icon:
  On Error GoTo Reset_Key
  
  If PixelsXY = 16 Then
    SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
  ElseIf PixelsXY = 32 Then
    SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
  Else
    MsgBox "Icon is an invalid size", vbCritical
    Exit Function
  End If
    
  If SmallIcon <> 0 Then
    With PictureBox
      .Height = PixelsXY
      .Width = PixelsXY
      .ScaleHeight = PixelsXY
      .ScaleWidth = PixelsXY
      .Picture = LoadPicture("")
      .AutoRedraw = True
      
      SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    
    IconIndex = AddtoImageList.ListImages.Count + 1
    Set NewImage = AddtoImageList.ListImages.Add(IconIndex, iKey, PictureBox.Image)
    ExtractIcon = IconIndex
  End If
  Exit Function
  
Reset_Key:
  iKey = ""
  Resume
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 1 Then key_Shift = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then key_Shift = False
End Sub

Private Sub lvDir_AfterLabelEdit(Cancel As Integer, NewString As String)
  'MsgBox NewString
  'Set thisFile = lvDir.SelectedItem.Tag
End Sub

Private Sub lvDir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lvMousePos.x = ScaleX(x, vbTwips, vbPixels)
  lvMousePos.y = ScaleY(y, vbTwips, vbPixels)
End Sub

Private Sub mnuPreviewPicture_Click()
  PreviewAs pvpicture
End Sub

Private Sub mnuPreviewText_Click()
  PreviewAs pvText
End Sub

Private Sub mnuPreviewVideo_Click()
  PreviewAs pvVideo
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim DeleteType As String
  
  On Error Resume Next
  Select Case Button.Key
    Case "Up"
      If Left(Dir1.Tag, 2) = "Ok" Then
        Dir1.Path = Dir1.List(-2)
        SetTreePath Dir1.Path
      End If
    Case "Delete"
      DeleteType = IIf(lvDir.SelectedItem.SubItems(2) = "File Folder", "Folder", "File")
      
      If MsgBox("Are you sure you want to delete " & IIf(DeleteType = "Folder", "the " & LCase(DeleteType) & " ", "") & "'" & lvDir.SelectedItem.Text & "' ?", vbQuestion + vbYesNo, "Confirm " & DeleteType & " Delete") = vbYes Then
        If LCase(DeleteType) = "folder" Then
          fSys.DeleteFolder lvDir.SelectedItem.Key
        Else
          fSys.DeleteFile lvDir.SelectedItem.Tag
        End If
        lvDir.ListItems.Remove lvDir.SelectedItem.Index
      End If
    Case "Refresh"
      Dim PathNow As String
      Drive1.Refresh
      File1.Refresh
      File1.Refresh
            
      PathNow = Dir1.Path
      BuildDriveList
      SetTreePath PathNow
      RefreshLV
  End Select
End Sub

Private Sub BuildDriveList()
  Dim TreePath As String
  Dim TreeIcon As Long
  'Dim DriveType
  
  tvDir.Nodes.Clear
  For i = 0 To Drive1.ListCount - 1
    TreePath = UCase(Left(Drive1.List(i), 1)) & ":\"
    'DriveType = GetDriveType(Drive1.List(i))
    TreeIcon = ExtractIcon(TreePath, ImgSmall, picBuffer, 16, TreePath)
    
    tvDir.Nodes.Add , , TreePath, UCase(Drive1.List(i)), TreeIcon
    tvDir.Nodes.Add TreePath, tvwChild, ""
  Next
End Sub

Private Sub cmdClosePreview_Click()
  cmdClosePreview.Visible = False
  rtfText.Visible = False
  pPicture.Visible = False
  lvDir.Visible = True
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
  RefreshLV
End Sub

Private Sub Drive1_Change()

  On Error GoTo ErrorHandler

  Dir1.Path = Drive1.Drive
  File1.Path = Dir1.Path
  Exit Sub
    
ErrorHandler:
  Drive1.Drive = "c:"
  Dir1.Path = "c:"
  
End Sub

Private Sub Form_Load()
  'Call InitCommonControls

  Set fSys = CreateObject("Scripting.FileSystemObject")
  Left = GetSetting(App.Title, "Settings", "mLeft", 1000)
  Top = GetSetting(App.Title, "Settings", "mTop", 1000)
  Width = GetSetting(App.Title, "Settings", "mWidth", 10000)
  Height = GetSetting(App.Title, "Settings", "mHeight", 7000)
  
  BuildDriveList
  
  'On Error Resume Next
  Dim c As String: c = Command
  If c = "" Then c = "c:\program files\"
  
  If c <> "" Then
    SetTreePath c
  End If
End Sub

Private Function GetFName(PathFile_Name As String)
  Dim j As Integer
  For j = 1 To Len(PathFile_Name)
    If InStr(Right(PathFile_Name, j), "\") = 1 Then Exit For
  Next j
  GetFName = Right(PathFile_Name, j - 1)
End Function

Private Sub RefreshLV(Optional ShowFolderSize As Boolean)
  Dim LItem As ListItem
  Dim FName As String 'Correct Path and filename
  Dim r As Integer
  
  Caption = "Explorer - " & Dir1.Path
  
  If IsMissing(ShowFolderSize) Then ShowFolderSize = False
  
  If Not lvDir.Visible Then cmdClosePreview_Click
  
  tvDir.Enabled = False
  lvDir.ListItems.Clear
  
  For i = 1 To Dir1.ListCount  'Add Folders to ListView
    Set thisFolder = fSys.GetFolder(Dir1.List(i - 1))
    r = ExtractIcon(thisFolder.Path, ImgSmall, picBuffer, 16, thisFolder.Type)
    Set LItem = lvDir.ListItems.Add(, Dir1.List(i - 1), thisFolder.Name, , r)
    If ShowFolderSize Then
      LItem.SubItems(1) = Format(thisFolder.Size / 1024, "###,##0") & " KB"
      'If i Mod 1 = 0 Then DoEvents 'Update the display
    Else
      LItem.SubItems(1) = " "
      'If i Mod 10 = 0 Then DoEvents 'Update the display
    End If
    LItem.SubItems(2) = thisFolder.Type
    DoEvents
  Next i
  
  For i = 1 To File1.ListCount  'Add Files to ListView
    FName = File1.Path & IIf(Right(File1.Path, 1) <> "\", "\", "") & File1.List(i - 1)
    Set thisFile = fSys.GetFile(FName)
    r = ExtractIcon(FName, ImgSmall, picBuffer, 16, thisFile.Type)
    Set LItem = lvDir.ListItems.Add(, , thisFile.Name, , r)
    LItem.Tag = FName
    LItem.SubItems(1) = Format(thisFile.Size / 1024, "###,##0") & " KB"
    LItem.SubItems(2) = thisFile.Type
    'If i Mod 50 = 0 Then DoEvents
    DoEvents
  Next i
  tvDir.Enabled = True
End Sub

Function SetTreePath(MyPath As String)
  On Error GoTo ErrorTreeView
    
  Dim SubDirNum As Integer, Dummy As Integer, NextSlash As Integer
  Dim MyFolder(0 To 20) As String
  Dim MyDir
  
  i = 0
  NextSlash = 1
  DummyDir.Path = MyPath
    
  If Right(DummyDir.Path, 1) <> "\" Then 'if the path isnt the root
    Do
      MyFolder(i) = Left(MyPath, InStr(NextSlash, MyPath, "\") - 1)
      i = i + 1
      NextSlash = InStr(NextSlash, MyPath, "\", 0) + 1
    Loop Until InStr(NextSlash, MyPath, "\", 0) = 0
    MyFolder(0) = UCase(MyFolder(0)) & "\"
    MyFolder(i) = MyPath
    SubDirNum = i
    For Dummy = 0 To SubDirNum - 1 'change the +1 to 0 OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
      tvDir.Nodes(MyFolder(Dummy)).Expanded = True
    Next Dummy
  End If
    
  Dir1.Path = MyPath
  MyDir = Dir1.Path
  If Right(MyDir, 1) <> "\" Then MyDir = MyDir & "\"
  tvDir.Nodes(Dir1.Path).Selected = True
  If Right(Dir1.Path, 1) = "\" Then
    Dir1_Change 'to avoid error when dir is root and the picutres not shown
  Else
    tvDir.Nodes(Dir1.Path).SelectedImage = 6
  End If

ErrorTreeView:

End Function

Private Sub Form_Resize()
  On Error Resume Next
  If Width < 3000 Then Width = 3000
  SizeControls imgSplitter.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "mLeft", Left
    SaveSetting App.Title, "Settings", "mTop", Top
    SaveSetting App.Title, "Settings", "mWidth", Width
    SaveSetting App.Title, "Settings", "mHeight", Height
  End If
  End
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  x = ScaleX(x, vbTwips, vbPixels)
  If mbMoving Then
    sglPos = x + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > ScaleWidth - sglSplitLimit Then
      picSplitter.Left = ScaleWidth - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Private Sub lvDir_DblClick()
  Dim rc As RECT
  Dim ItemNum As Integer
  
  For i = 1 To lvDir.ListItems.Count
    rc.Left = LVIR_BOUNDS
    SendMessage lvDir.hwnd, LVM_GETITEMRECT, i - 1, rc
    With lvMousePos
      If .x >= rc.Left And .x <= rc.Right And .y >= rc.Top And .y <= rc.Bottom Then
        ItemNum = i
      End If
    End With
  Next i

  If ItemNum <> 0 Then
    Dim lvItem As ListItem
  
    Set lvItem = lvDir.ListItems(ItemNum)
    
    SetTreePath lvItem.Key
    If lvItem.Tag <> "" Then mnuPreview_Click
  End If
End Sub

Private Sub lvDir_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With lvDir
    '.SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    .SortKey = ColumnHeader.Index - 1
    .Sorted = True
    .Sorted = False
  End With
End Sub

Private Sub lvDir_ItemClick(ByVal Item As MSComctlLib.ListItem)
  On Error GoTo ShowError
  
  PopUpMnu = False
  If lvButton = 1 Then
    If Item.Key <> "" And lvMouseDown And key_Shift Then
      Set thisFolder = fSys.GetFolder(Item.Key)
      Item.SubItems(1) = Format(thisFolder.Size / 1024, "###,##0") & " KB"
    End If
  ElseIf lvButton = 2 Then
    If Item.Key = "" Then 'File has been clicked
      PopUpMnu = True
    Else
      RefreshLV True
    End If
  End If
  Exit Sub
  
ShowError: MsgBox Error, vbCritical
End Sub

Private Sub lvDir_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  Select Case KeyAscii
    Case 8: tbToolBar_ButtonClick tbToolBar.Buttons("Up")
    Case 13 And Right(lvDir.SelectedItem.SubItems(2), 6) = "Folder": SetTreePath lvDir.SelectedItem.Key
    Case Else ': MsgBox KeyAscii
  End Select
End Sub

Private Sub lvDir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  lvButton = Button
  lvMouseDown = True
End Sub

Private Sub lvDir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    If PopUpMnu Then PopupMenu mnuPopUp
  End If
  PopUpMnu = False
  lvMouseDown = False
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuNewFolder_Click()
  'Set thisFolder = fSys.CreateFolder(Dir1.Path & IIf(Right(Dir1.Path, 1) <> "\", "\", "") & "Tims New Folder1")
End Sub

Private Sub mnuPreview_Click()
  Dim fType As String
  
  fType = Right(lvDir.SelectedItem.Text, 3)
  
  Select Case LCase(fType)
    Case "txt", "rtf", "dat", "bak", "frm", "vbp", "bas"
      mnuPreviewText_Click
    Case "bmp", "jpg", "gif", "ico"
      mnuPreviewPicture_Click
    Case "mp3", "wma", "avi"
      mnuPreviewVideo_Click
    Case "exe"
      On Error Resume Next
      Shell lvDir.SelectedItem.Tag, vbNormalFocus
  End Select
End Sub

Private Sub PreviewAs(pType As PreviewType)
  
  On Error GoTo ShowError
  
  lvDir.Visible = False
  cmdClosePreview.Visible = True
  
  Select Case pType
    Case pvText
      rtfText.Visible = True
      rtfText.LoadFile lvDir.SelectedItem.Tag
    Case pvpicture
      pPicture.Visible = True
      pPicture.Picture = LoadPicture(lvDir.SelectedItem.Tag)
    Case pvVideo: MsgBox "Feature not currently working!!", vbExclamation
    Case Else: MsgBox "No Preview Available", vbInformation
  End Select
  Exit Sub

ShowError:
  
  MsgBox "That Is Impossible!!", vbExclamation
End Sub

Private Sub SizeControls(x As Single)
  'On Error Resume Next
  
  If x < 100 Then x = 100
  If x > (ScaleWidth - 100) Then x = ScaleWidth - 100
  tvDir.Width = x
  imgSplitter.Left = x
  lvDir.Left = x + 4
  lvDir.Width = ScaleWidth - (tvDir.Width + 4)
  
  tvDir.Top = tbToolBar.Height
  lvDir.Top = tvDir.Top
  
  tvDir.Height = ScaleHeight - tbToolBar.Height
  lvDir.Height = tvDir.Height
  
  imgSplitter.Top = tvDir.Top
  imgSplitter.Height = tvDir.Height
  
  cmdClosePreview.Left = lvDir.Left
  pPicture.Left = lvDir.Left
  
  With lvDir
    rtfText.Move .Left, rtfText.Top, .Width, .Height - cmdClosePreview.Top
  End With
End Sub

Private Sub tvDir_Expand(ByVal Node As MSComctlLib.Node)
  Dim CurrentPath As String, FolderName As String
  Dim r As Integer
  
  On Error GoTo ErrorTreeView

  DummyDir.Path = Node.Key
    
  If Node.Child.Text = "" Then
    tvDir.Nodes.Remove Node.Child.Index
    For i = 0 To DummyDir.ListCount - 1
      FolderName = Mid(DummyDir.List(i), Len(DummyDir.Path) + 2)
      If Len(DummyDir.Path) = 3 Then FolderName = Mid(DummyDir.List(i), Len(DummyDir.Path) + 1)
      
      r = ExtractIcon(DummyDir.List(i), ImgSmall, picBuffer, 16, "File Folder")
      
      tvDir.Nodes.Add DummyDir.Path, tvwChild, DummyDir.List(i), FolderName, r
      CheckForChildDir.Path = DummyDir.List(i) 'checking for childs
      If CheckForChildDir.ListCount > 0 Then
        tvDir.Nodes.Add DummyDir.List(i), tvwChild, ""
        tvDir.Nodes(DummyDir.List(i)).ExpandedImage = r
      End If
    Next i
  End If
  Dir1.Tag = "Ok 2 Up"
    
ErrorTreeView:
  On Error Resume Next
  If Node.Child.Text = "" Then tvDir.Nodes(Node.Index).Expanded = False
  
End Sub

Private Sub tvDir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  MouseButtonPressed = Button
End Sub

Private Sub tvDir_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo NodeClickError

  If MouseButtonPressed = 2 Then
    'here you can put code for the TreeView pup-up menu
  Else
    Dir1.Path = Node.Key
    If Right(Node.Key, 1) <> "\" Then Node.SelectedImage = 6
  End If

NodeClickError:

End Sub

  
'  Dim rc As RECT, i As Long
'  For i = 1 To 1 'ListView1.ListItems.Count
'      rc.Left = LVIR_BOUNDS
'      SendMessage ListView1.hWnd, LVM_GETITEMRECT, i - 1, rc
'      TTBlnGreen.AjustItemRect ListView1.hWnd, i - 1, rc.Left, rc.Top, rc.Right, rc.Bottom
'  Next i
'  Label1 = ScaleX(X, vbTwips, vbPixels) & " " & ScaleY(Y, vbTwips, vbPixels)
'  Label2 = "X " & rc.Left & " " & rc.Right & " Y " & rc.Top & " " & rc.Bottom

