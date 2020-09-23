Attribute VB_Name = "mMain"
Option Explicit

Private Type BrowseInfo
    hwndOwner                                    As Long
    pIDLRoot                                     As Long
    pszDisplayName                               As Long
    lpszTitle                                    As Long
    ulFlags                                      As Long
    lpfnCallback                                 As Long
    lParam                                       As Long
    iImage                                       As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                                   ByVal lpNewFileName As String, _
                                                                   ByVal bFailIfExists As Long) As Long

Public Function Get_Path() As String
'/* get scan path dialog

Dim lpIDList    As Long
Dim szTitle     As String
Dim tBrowseInfo As BrowseInfo
Dim sBuffer     As String

On Error Resume Next

    szTitle = "Select a Directory to Scan: "

    With tBrowseInfo
        .hwndOwner = frmMain.hWnd
        .lpszTitle = lstrcat(szTitle, vbNullString)
        .ulFlags = 1 + 2
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If lpIDList Then
        sBuffer = Space$(260)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Get_Path = sBuffer + Chr$(92)
    End If

On Error GoTo 0

End Function
