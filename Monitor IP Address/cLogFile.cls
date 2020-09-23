VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MAX_PATH = 260

'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarMaxSize As Long 'local copy
Private mvarDateFormat As String 'local copy
Private mvarTimeFormat As String 'local copy
Private mvarUseDate As Boolean 'local copy
Private mvarUseTime As Boolean 'local copy
Private mvarSpaceBetweenLogs As Integer 'local copy
Private mvarHeader As String 'local copy
Private mvarLineSpaceFormat As String 'local copy

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Function GetLogFile(Optional FileName As String)

Dim iNextFree As Integer
Dim strMsg As String
    
    If FileName = "" Then
        FileName = mvarFileName
    End If
    
    If FileName = "" Then Exit Function
    If Len(Dir(FileName, vbDirectory)) = 0 Then Exit Function
    
    iNextFree = FreeFile
    Open FileName For Binary As iNextFree
        strMsg = Input(FileLen(mvarFileName), #iNextFree)
    Close iNextFree
    
    GetLogFile = strMsg
    
End Function


Public Sub KillLogFile()
    If Len(mvarFileName) = 0 Then Exit Sub
    If Len(Dir(mvarFileName, vbDirectory)) <> 0 Then Kill mvarFileName
End Sub

Public Property Let LineSpaceFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LineSpaceFormat = 5
    mvarLineSpaceFormat = vData
End Property


Public Property Get LineSpaceFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LineSpaceFormat
    LineSpaceFormat = mvarLineSpaceFormat
End Property



Public Property Let Header(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Header = 5
    mvarHeader = vData
End Property


Public Property Get Header() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Header
    Header = mvarHeader
End Property




Public Sub AppendLog(Message As String)

Dim iNextFree As Integer
Dim strMsg As String

    If mvarFileName = "" Then Exit Sub
    If Me.DoesLogExist(mvarFileName) = False Then
        CreateLog
    End If
    
    
    strMsg = mvarLineSpaceFormat
    
    If mvarUseDate = True Then
        strMsg = strMsg & " " & Format(Now, mvarDateFormat)
    End If
    
    If mvarUseTime = True Then
        strMsg = strMsg & " " & Format(Now, mvarTimeFormat)
    End If
    
    strMsg = strMsg & " -" & vbTab & Message
    
    iNextFree = FreeFile
    Open mvarFileName For Append As iNextFree
        Print #iNextFree, strMsg
    Close iNextFree
    

End Sub
Property Get WindowsDirectory() As String
  ' Returns : The Windows directory
  ' Source: Total VB SourceBook 6

  Dim strWindowsDirectory As String * MAX_PATH
  
  On Error GoTo PROC_ERR
  
  GetWindowsDirectory strWindowsDirectory, MAX_PATH
  
  WindowsDirectory = TrimNulls(strWindowsDirectory)

PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WindowsDirectory"
  Resume PROC_EXIT

End Property
Public Sub CreateLog(Optional Header As String, Optional FileName As String)

Dim iNextFree As Integer

    If FileName = "" Then
        FileName = mvarFileName
    End If
    
    If FileName = "" Then
        mvarFileName = Me.CreateTempFile("win")
        FileName = mvarFileName
    End If
    
    If Len(Header) = 0 Then
        Header = mvarHeader
    End If
    
    iNextFree = FreeFile
    Open FileName For Output As iNextFree
        If Len(Header) <> 0 Then
            Print #iNextFree, Header
        End If
    Close iNextFree

End Sub
Public Function DoesLogExist(Optional FileName As String) As Boolean

    DoesLogExist = False

    If FileName = "" Then
        FileName = mvarFileName
    End If
    
    If FileName = "" Then Exit Function
    
    If Len(Dir(FileName, vbDirectory)) <> 0 Then
        DoesLogExist = True
    End If
    
End Function


Public Property Let UseTime(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UseTime = 5
    mvarUseTime = vData
End Property


Public Property Get UseTime() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UseTime
    UseTime = mvarUseTime
End Property



Public Property Let UseDate(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UseDate = 5
    mvarUseDate = vData
End Property


Public Property Get UseDate() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UseDate
    UseDate = mvarUseDate
End Property



Public Property Let TimeFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TimeFormat = 5
    If Len(vData) = 0 Then vData = "hh:mm:ss"
    mvarTimeFormat = vData
End Property


Public Property Get TimeFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TimeFormat
    TimeFormat = mvarTimeFormat
End Property



Public Property Let DateFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateFormat = 5
    If Len(vData) = 0 Then vData = "mm/dd/yyyy"
    mvarDateFormat = vData
End Property


Public Property Get DateFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateFormat
    DateFormat = mvarDateFormat
End Property



Public Function ClearLogFile()
    If Len(Dir(mvarFileName, vbDirectory)) <> 0 Then Kill mvarFileName
End Function


Public Property Let MaxSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxSize = 5
    mvarMaxSize = vData
End Property


Public Property Get MaxSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxSize
    MaxSize = mvarMaxSize
End Property



Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



Public Function CreateTempFile(strPrefix As String) As String
  ' Returns    : A temporary file name based on the value of strPrefix.
  ' Source: Total VB SourceBook 6
  '
  Dim strTemp As String
  Dim lngRet As Long
  Dim strTempPath As String
    
  On Error GoTo PROC_ERR
  
  strTempPath = Space$(255)
  lngRet = GetTempPath(Len(strTempPath), strTempPath)
  
  strTemp = Space$(255)
  lngRet = GetTempFileName(strTempPath, strPrefix, 1, ByVal strTemp)
  
  strTemp = TrimNulls(strTemp)
  
  If Len(Dir(strTemp, vbDirectory)) <> 0 Then Kill strTemp
  
  CreateTempFile = strTemp

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CreateTempFile"
  Resume PROC_EXIT

End Function

Private Function TrimNulls(ByVal strIn As String) As String
  ' Comments  : Returns the passed string terminated at the first null
  ' Parameters: strIn - Value to parse
  ' Returns   : Parsed string
  ' Source    : Total VB SourceBook 6
  '
  Dim intPos As Integer
  
  On Error GoTo PROC_ERR
    
  intPos = InStr(strIn, vbNullChar)
  
  If intPos = 0 Then
    ' No nulls in the string, just return it as is
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      TrimNulls = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      TrimNulls = Left$(strIn, intPos - 1)
    End If
  End If
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TrimNulls"
  Resume PROC_EXIT
    
End Function
Private Sub Class_Initialize()

    mvarFileName = CreateTempFile("win")
    mvarMaxSize = 5242880   ' 5 MB
    mvarDateFormat = "mm/dd/yyyy"
    mvarTimeFormat = "hh:mm:ss"
    mvarUseDate = True
    mvarUseTime = True
    mvarSpaceBetweenLogs = 0

End Sub
