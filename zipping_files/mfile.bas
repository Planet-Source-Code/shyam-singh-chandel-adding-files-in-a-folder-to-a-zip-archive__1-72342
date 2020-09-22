Attribute VB_Name = "mfile"
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'////                                                                    /////
'////     Developer: Shyam Singh Chandel                                 /////
'////     shyamschandel@ rediffmail.com                                  /////
'////     shyamschandel@developerssourcecode.com                         /////
'////     Programmer, System, Hardware and Electronic Engineer           /////
'////     URL http://www.developerssourcecode.com                        /////
'////                                                                    /////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////

Option Explicit



Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Const MAX_PATH = 260
Private Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_STATUSTEXT = &H4
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Const FileAttributeArchive = &H20, FileAttributeReadonly = &H1
Private Const FileAttributeSystem = &H4, FileAttributeHidden = &H2
Private Const FileAttributeDirectory = &H10
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Function BrowseForFolder(Prompt As String) As String
  Dim n As Integer
  Dim IDList As Long
  Dim Result As Long
  Dim ThePath As String
  Dim BI As BrowseInfo
  With BI
    .hWndOwner = GetActiveWindow()
    .lpszTitle = lstrcat(Prompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  IDList = SHBrowseForFolder(BI)
  If IDList Then
    ThePath = String$(MAX_PATH, 0)
    Result = SHGetPathFromIDList(IDList, ThePath)
    Call CoTaskMemFree(IDList)
    n = InStr(ThePath, vbNullChar)
    If n Then ThePath = Left$(ThePath, n - 1)
  End If
  BrowseForFolder = ThePath
End Function

Public Function CreateDirectory(ByVal psPath As String) As Integer
    Dim nCallDrive As String
    Dim pos1, pos2, pos3 As Integer
    Dim sDrive As String
    Dim sTemp As String
    On Error GoTo CreateDirErr
    sDrive = Left$(LTrim$(psPath), 3)
    nCallDrive = DriveExists(sDrive)
    If nCallDrive = "" Then
        MsgBox "Drive Does not exist ", 16
        Exit Function
    End If
    If Right$(Trim$(psPath), 1) = "\" Then
      psPath = Left$(Trim$(psPath), Len(Trim$(psPath)) - 1)
    End If
    If Dir$(psPath, 16) = "" Then
        pos1 = 3
        ChDrive sDrive
        ChDir "\"
        Do
            pos2 = pos1
            pos1 = InStr(pos2 + 1, psPath, "\")
            If pos1 > 0 Then
                sTemp = Left$(psPath, pos1 - 1)
            Else
                sTemp = psPath
            End If
            MkDir sTemp
            ChDir sTemp
        Loop While pos1 > 0
    Else
        If InStr(Trim$(sDrive), ":") <> 0 Then
          ChDrive sDrive
        End If
        ChDir psPath
    End If
    psPath = LCase$(psPath)
    CreateDirectory = True
    Exit Function

CreateDirErr:
    If Err = 75 Or 76 Then Resume Next
    CreateDirectory = False
    Exit Function
    Resume
End Function

Function DriveExists(ByVal drive As String) As String
    DriveExists = ""
    If Left$(drive, 2) <> "\\" Then
        DriveExists = Dir$(drive, 16)
    Else
        DriveExists = "\\"
    End If

End Function


Public Function IsFolderExists(ByVal FolderName As String) As Boolean
Dim FS
Set FS = CreateObject("Scripting.FileSystemObject")
IsFolderExists = FS.FolderExists(FolderName)
End Function

Public Function bFileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
i = Len(Dir$(sFileName))
bFileExists = IIf(Err Or i = 0, False, True)
If Trim(sFileName) = vbNullString Then bFileExists = False
End Function
Public Function ReadINI(sSection As String, sKeyName As String, sINIFileName _
As String) As String
Dim sRet As String: sRet = String(255, Chr(0))
ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, vbNullString, sRet, Len(sRet), sINIFileName))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sValueData _
As String, sINIFileName As String) As Boolean
Call WritePrivateProfileString(sSection, sKeyName, sValueData, _
sINIFileName)
WriteINI = (Err.Number = 0)
End Function
Public Sub FileAttribHide(ByVal FileName As String)
Dim vResult As Long
vResult = SetFileAttributes(FileName, FileAttributeHidden + FileAttributeSystem)
End Sub
Public Sub deltrees(ByVal sDir As Variant)
    Dim FSO, FS
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FS = FSO.DeleteFolder(sDir, True)
End Sub


