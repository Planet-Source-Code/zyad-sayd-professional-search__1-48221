Attribute VB_Name = "Module1"
Public NewFolder As String
Public D As Integer

'And This one for the icons.
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal rGetIcon As Long) As Long

Public FilesNum As Long
'Move File To RecycleBin
Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As Long
End Type

Public Declare Function SHFileOperation Lib _
"shell32.dll" Alias "SHFileOperationA" (lpFileOp _
As SHFILEOPSTRUCT) As Long



'Copy Files or Folder
Public Const FOF_ALLOWUNDO = &H40
'File Operation constants
Public Enum FO_OPS
    FO_MOVE = &H1
    FO_COPY = &H2
    FO_DELETE = &H3
    FO_RENAME = &H4
End Enum


'Folders Show
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

'To open the Appl.
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters  As String
    lpDirectory   As String
    nShow As Long
    hInstApp As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type

