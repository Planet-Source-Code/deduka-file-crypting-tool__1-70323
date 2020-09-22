Attribute VB_Name = "cod_BrowseForFolder"
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'Browsing type.
Public Enum BrowseType
    BrowseForFolders = &H1
    BrowseForComputers = &H1000
    BrowseForPrinters = &H2000
    BrowseForEverything = &H4000
End Enum

'Folder Type
Public Enum FolderType
    CSIDL_BITBUCKET = 10
    CSIDL_CONTROLS = 3
    CSIDL_DESKTOP = 0
    CSIDL_DRIVES = 17
    CSIDL_FONTS = 20
    CSIDL_NETHOOD = 18
    CSIDL_NETWORK = 19
    CSIDL_PERSONAL = 5
    CSIDL_PRINTERS = 4
    CSIDL_PROGRAMS = 2
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_STARTMENU = 11
End Enum

Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, listID As Long) As Long
'dodo:
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Private Const WM_USER                   As Long = &H400
Private Const BFFM_INITIALIZED          As Long = 1
Private Const BFFM_SELCHANGED           As Long = 2
Private Const BFFM_SETSTATUSTEXT        As Long = WM_USER + 100
Private Const BFFM_ENABLEOK             As Long = WM_USER + 101
Private Const BFFM_SETSELECTION         As Long = WM_USER + 102

Dim InitialFolder As String
Dim BInfo As BrowseInfo


' This function compensates for the fact that the AddressOf operator
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function
' This function is the callback function for the dialog box.
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  Dim retval As Long  ' return value
  If uMsg = BFFM_INITIALIZED Then
    retval = SendMessage(hwnd, BFFM_SETSELECTION, ByVal CLng(1), ByVal InitialFolder)
  End If
  BrowseCallbackProc = 0
End Function


Public Function BrowseFolders(hwndOwner As Long, sMessage As String, Browse As BrowseType, Optional ByVal RootFolder As FolderType, Optional InitFolder As String) As String
  Dim Nullpos As Integer
  Dim lpIDList As Long
  Dim res As Long
  Dim sPath As String
  Dim RootID As Long

  SHGetSpecialFolderLocation hwndOwner, RootFolder, RootID
  BInfo.hwndOwner = hwndOwner
  BInfo.lpszTitle = lstrcat(sMessage, "")
  BInfo.ulFlags = Browse
  If RootID <> 0 Then BInfo.pIDLRoot = RootID
  InitialFolder = InitFolder
  BInfo.lpfnCallback = DummyFunc(AddressOf BrowseCallbackProc)

  lpIDList = SHBrowseForFolder(BInfo)
  If lpIDList <> 0 Then
    sPath = String(MAX_PATH, 0)
    res = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    Nullpos = InStr(sPath, vbNullChar)
    If Nullpos <> 0 Then
      sPath = left(sPath, Nullpos - 1)
    End If
  End If
    
  If sPath = "" Then
    BrowseFolders = ""
  Else
    BrowseFolders = sPath & IIf(Right(sPath, 1) = "\", "", "\")
  End If
End Function

