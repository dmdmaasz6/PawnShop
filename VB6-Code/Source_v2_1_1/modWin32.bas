Attribute VB_Name = "modWin32"
Option Explicit
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
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

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hWnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             lParam As Any) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
            (ByVal lpFileName As String, _
             lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" _
            (ByVal hFindFile As Long) As Long

Public Function FileExists(sSource As String) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)
   
End Function

Public Function FolderExists(sFolder As String) As Boolean

   Dim hFile As Long
   Dim WFD As WIN32_FIND_DATA
   
   sFolder = UnQualifyPath(sFolder)
   hFile = FindFirstFile(sFolder, WFD)

   FolderExists = (hFile <> INVALID_HANDLE_VALUE) And _
                  (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
   
   Call FindClose(hFile)
   
End Function

Private Function UnQualifyPath(ByVal sFolder As String) As String

   sFolder = Trim$(sFolder)
   
   If Right$(sFolder, 1) = "\" Then
      UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
   Else
      UnQualifyPath = sFolder
   End If
   
End Function

