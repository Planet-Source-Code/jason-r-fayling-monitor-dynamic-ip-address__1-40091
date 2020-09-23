Attribute VB_Name = "modMonitorIPAddress"
Option Explicit

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const LVM_FIRST = &H1000
    Private Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
    Private Const LVM_SETITEMSTATE = LVM_FIRST + 43
    Private Const LVM_GETITEMSTATE = LVM_FIRST + 44
    Private Const LVIS_STATEIMAGEMASK = &HF000
    Private Const LVM_GETITEM = LVM_FIRST + 5 '75 for unicode?
    Private Const LVIF_STATE = &H8
    
    Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
    Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
    Private Const LVS_EX_FULLROWSELECT = &H20
    Private Const WM_SETREDRAW = &HB
    Private Const LVS_EX_GRIDLINES = &H1
    Private Const LVS_EX_SUBITEMIMAGES = &H2
    Private Const LVS_EX_CHECKBOXES = &H4
    Private Const LVS_EX_TRACKSELECT = &H8
    Private Const LVS_EX_HEADERDRAGDROP = &H10
    
    Private Const LVSCW_AUTOSIZE = -1
    Private Const LVSCW_AUTOSIZE_USEHEADER = -2
    
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Private Const SND_APPLICATION = &H80         '  look for application specific association
    Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
    Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
    Private Const SND_ASYNC = &H1         '  play asynchronously
    Private Const SND_FILENAME = &H20000     '  name is a file name
    Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
    Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
    Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
    Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
    Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
    Private Const SND_PURGE = &H40               '  purge non-static events for task
    Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
    Private Const SND_SYNC = &H0         '  play synchronously (default)
Public Sub AutoSizeColumns(m_ListView As Object)
  ' Comments  : Sizes each column in the listview control to fit
  '             the widest data in each column
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim intColumn As Integer
  
  On Error GoTo PROC_ERR
  
  For intColumn = 0 To m_ListView.ColumnHeaders.Count - 1
    SendMessageLong _
      m_ListView.hwnd, _
      LVM_SETCOLUMNWIDTH, _
      intColumn, _
      LVSCW_AUTOSIZE_USEHEADER
  Next intColumn

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AutoSizeColumns"
  Resume PROC_EXIT

End Sub
Public Sub PlayNotifySound(ByVal sFilename As String)

    If DoesFileExist(sFilename) = False Then Exit Sub
    PlaySound sFilename, ByVal 0&, SND_FILENAME Or SND_ASYNC

End Sub
Public Sub UpdateHOSTS(ByVal sHOSTSName As String, ByVal sIPAddress As String)

On Error GoTo ErrorHandle

Dim sPath As String
Dim sFile As String
Dim sBuffer As String
Dim sNewBuffer As String
Dim i As Long
Dim j As Long
Dim k As Long

    sPath = String(255, " ")
    GetWindowsDirectory sPath, Len(sPath)
    
    sPath = Trim(sPath)
    sPath = Left(sPath, Len(sPath) - 1)
    
    sFile = sPath & "\HOSTS"
    If Len(Dir(sFile, vbNormal)) = 0 Then
        sFile = sPath & "\System32\Drivers\etc\HOSTS"
        If Len(Dir(sFile, vbNormal)) = 0 Then
            Exit Sub
        End If
    End If
    
    sBuffer = OpenFileAsString(sFile)
    j = InStr(1, UCase(sBuffer), UCase(sHOSTSName), vbBinaryCompare)
    If j <> 0 Then
        i = InStrRev(sBuffer, vbCrLf, j)
        If i <> 0 Then
            i = i - 1
            sNewBuffer = Mid(sBuffer, 1, i) & vbCrLf
            sNewBuffer = sNewBuffer & sIPAddress & " "
            sNewBuffer = sNewBuffer & Mid(sBuffer, j)
        End If
    Else
        sNewBuffer = sBuffer & vbCrLf & sIPAddress & " " & sHOSTSName
    End If
    
    WriteStringAsFile sNewBuffer, sFile

ClearVariables:
    Exit Sub
    
ErrorHandle:
    GoTo ClearVariables

End Sub
Public Function OpenFileAsString(FileName As String) As String

On Error GoTo ErrorCode

Dim iNextFree As Integer
Dim bolDidOpen As Boolean

    If DoesFileExist(FileName) = False Then GoTo ClearVariables
    iNextFree = FreeFile
    
    Open FileName For Binary As iNextFree
        bolDidOpen = True
        OpenFileAsString = Input(FileLen(FileName), #iNextFree)

ClearVariables:
    If bolDidOpen = True Then Close iNextFree
    Exit Function
    
ErrorCode:
    OpenFileAsString = ""
    GoTo ClearVariables

End Function

Public Function DoesFileExist(FileName As String) As Boolean

On Error GoTo ErrorCode
    
    DoesFileExist = False
        If Len(FileName) = 0 Then GoTo ClearVariables
        If Len(Dir(FileName, vbNormal)) = 0 Then GoTo ClearVariables
    DoesFileExist = True

ClearVariables:
    Exit Function
ErrorCode:
    DoesFileExist = False
    GoTo ClearVariables

End Function
Public Function WriteStringAsFile(Buffer As String, FileName As String) As Boolean

On Error GoTo ErrorCode

Dim iNextFree As Integer
Dim bolDidOpen As Boolean

    WriteStringAsFile = False
    iNextFree = FreeFile
    
    Open FileName For Binary As iNextFree
        bolDidOpen = True
        Put #iNextFree, , Buffer
    WriteStringAsFile = True
    
ClearVariables:
    If bolDidOpen = True Then Close iNextFree
    Exit Function
    
ErrorCode:
    WriteStringAsFile = False
    GoTo ClearVariables

End Function
