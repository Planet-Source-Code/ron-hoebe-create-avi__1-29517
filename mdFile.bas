Attribute VB_Name = "mdFile"
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
   Private Const DT_BOTTOM = &H8&
   Private Const DT_CENTER = &H1&
   Private Const DT_LEFT = &H0&
   Private Const DT_CALCRECT = &H400&
   Private Const DT_WORDBREAK = &H10&
   Private Const DT_VCENTER = &H4&
   Private Const DT_TOP = &H0&
   Private Const DT_TABSTOP = &H80&
   Private Const DT_SINGLELINE = &H20&
   Private Const DT_RIGHT = &H2&
   Private Const DT_NOCLIP = &H100&
   Private Const DT_INTERNAL = &H1000&
   Private Const DT_EXTERNALLEADING = &H200&
   Private Const DT_EXPANDTABS = &H40&
   Private Const DT_CHARSTREAM = 4&
   Private Const DT_NOPREFIX = &H800&
   Private Const DT_EDITCONTROL = &H2000&
   Private Const DT_PATH_ELLIPSIS = &H4000&
   Private Const DT_END_ELLIPSIS = &H8000&
   Private Const DT_MODIFYSTRING = &H10000
   Private Const DT_RTLREADING = &H20000
   Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" ( _
  ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long

Public Function CompactedPath( _
     ByVal sPath As String, _
     ByVal lMaxPixels As Long, _
     ByVal hDC As Long _
   ) As String
    Dim tR As RECT
   tR.right = lMaxPixels
   DrawText hDC, sPath, -1, tR, DT_PATH_ELLIPSIS Or DT_SINGLELINE Or DT_MODIFYSTRING
   CompactedPath = sPath
End Function

Public Function CompactedPathSh( _
     ByVal sPath As String, _
     ByVal lMaxPixels As Long, _
     ByVal hDC As Long _
   ) As String
    Dim lR As Long
    Dim iPos As Long
    lR = PathCompactPath(hDC, sPath, lMaxPixels)
    iPos = InStr(sPath, Chr$(0))
    If iPos <> 0 Then
      CompactedPathSh = left$(sPath, iPos - 1)
    Else
      CompactedPathSh = sPath
    End If
End Function

Public Function ReadINI(AppName As String, KeyName As String, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(AppName As String, KeyName As String, NewString As String, filename As String) As Integer
    WriteINI = WritePrivateProfileString(AppName, KeyName, NewString, filename)
End Function



' Creates a temporary (0 byte) file in the \TEMP directory
' and returns its name

Public Function GetTempFile(Optional Prefix As String) As String
    Dim TempFile As String
    Dim TempPath As String
    Const MAX_PATH = 260
    
    ' get the path of the \TEMP directory
    TempPath = Space$(MAX_PATH)
    GetTempPath Len(TempPath), TempPath
    ' trim off characters in excess
    TempPath = left$(TempPath, InStr(TempPath & vbNullChar, vbNullChar) - 1)
    
    ' get the name of a temporary file in that path, with a given prefix
    TempFile = Space$(MAX_PATH)
    GetTempFileName TempPath, Prefix, 0, TempFile
    GetTempFile = left$(TempFile, InStr(TempFile & vbNullChar, vbNullChar) - 1)

End Function

 
Function GetFileExtension(ByVal filename As String) As String
    Dim i As Long
    For i = Len(filename) To 1 Step -1
        Select Case Mid$(filename, i, 1)
            Case "."
                GetFileExtension = LCase$(Mid$(filename, i + 1))
                Exit For
            Case ":", "\"
                Exit For
        End Select
    Next
End Function

' Change the extension of a file name
' if the last argument is True, it adds the extension even if the file doesn't have one

Function ChangeFileExtension(filename As String, Extension As String, Optional AddIfMissing As Boolean) As String
    Dim i As Long
    For i = Len(filename) To 1 Step -1
        Select Case Mid$(filename, i, 1)
            Case "."
                ' we've found an extension, so replace it
                ChangeFileExtension = left$(filename, i) & Extension
                Exit Function
            Case ":", "\"
                Exit For
        End Select
    Next
    
    ' there is no extension
    If AddIfMissing Then
        ChangeFileExtension = filename & "." & Extension
    Else
        ChangeFileExtension = filename
    End If
End Function

' Retrieve a file's base name
' if the second argument is true, the result include the file's path

Function GetFileBaseName(filename As String, Optional ByVal IncludePath As Boolean) As String
    Dim i As Long, startPos As Long, endPos As Long
    
    startPos = 1
    
    For i = Len(filename) To 1 Step -1
        Select Case Mid$(filename, i, 1)
            Case "."
                ' we've found the extension
                If IncludePath Then
                    ' if we must return the path, we've done
                    GetFileBaseName = left$(filename, i - 1)
                    Exit Function
                End If
                ' else, just take note of where the extension begins
                If endPos = 0 Then endPos = i - 1
            Case ":", "\"
                If Not IncludePath Then startPos = i + 1
                Exit For
        End Select
    Next
    
    If endPos = 0 Then
        ' this file has no extension
        GetFileBaseName = Mid$(filename, startPos)
    Else
        GetFileBaseName = Mid$(filename, startPos, endPos - startPos + 1)
    End If
End Function

' Retrieve a file's path
'
' Note: trailing backslashes are never included in the result

Function GetFilePath(filename As String) As String
    Dim i As Long
    For i = Len(filename) To 1 Step -1
        Select Case Mid$(filename, i, 1)
            Case ":"
                ' colons are always included in the result
                GetFilePath = left$(filename, i)
                Exit For
            Case "\"
                ' backslash aren't included in the result
                GetFilePath = left$(filename, i - 1)
                Exit For
        End Select
    Next
End Function

' Append a backslash (or any character) at the end of a path
' if it isn't there already

Function AddBackslash(Path As String, Optional Char As String = "\") As String
    If right$(Path, 1) <> Char Then
        AddBackslash = Path & Char
    Else
        AddBackslash = Path
    End If
End Function

' Make a complete file name by assemblying its individual parts
' if Extension isn't omitted, it overwrites any extension held in BaseName

Function MakeFileName(Drive As String, Path As String, BaseName As String, _
    Optional Extension As String)

    ' add a trailing colon to the drive name, if needed
    MakeFileName = Drive & IIf(right$(Drive, 1) <> ":", ":", "")
    ' add a trailing backslash to the path, if needed
    MakeFileName = MakeFileName & Path & IIf(right$(Path, 1) <> "\", "\", "")
    
    If Len(Extension) = 0 Then
        ' no extension has been provided
        MakeFileName = MakeFileName & BaseName
    ElseIf InStr(BaseName, ".") = 0 Then
        ' the base name doesn't contain any extension
        MakeFileName = MakeFileName & BaseName & "." & Extension
    Else
        ' we need to drop the extension in the base name
        Dim i As Long
        For i = Len(BaseName) To 1 Step -1
            If Mid$(BaseName, i, 1) = "." Then
                MakeFileName = MakeFileName & left$(BaseName, i) & Extension
                Exit For
            End If
        Next
    End If
End Function
 
' Return True if a file exists

Function FileExists(filename As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(filename) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

' Return True if a directory exists
' (the directory name can also include a trailing backslash)

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

' returns the last occurrence of a substring
' The syntax is similar to InStr

Function InstrLast(ByVal start As Long, Source As String, search As String, _
    Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
    Do
        ' search the next occurrence
        start = InStr(start, Source, search, CompareMethod)
        If start = 0 Then Exit Do
        ' we found one
        InstrLast = start
        start = start + 1
    Loop
End Function



