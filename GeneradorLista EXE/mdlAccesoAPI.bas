Attribute VB_Name = "mdlAccesoAPI"
'----------------------------------------------------------------------------------
'Declaración de constantes necesaria para tranbar con la API FindFirstFile.
'----------------------------------------------------------------------------------
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_NO_MORE_FILES = 18&
'*

'-----------------------------------------------------------------------------
'Declaración de tipos necesarios para trabajar con la API FindFirstFile.
'-----------------------------------------------------------------------------
Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
'*


Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

