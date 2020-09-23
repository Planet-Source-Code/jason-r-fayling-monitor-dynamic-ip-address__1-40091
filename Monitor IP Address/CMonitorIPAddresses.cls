VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonitorIPAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents moTimer As XTimer
Attribute moTimer.VB_VarHelpID = -1
Private WithEvents moGetIPAddress As CGetIPAddress
Attribute moGetIPAddress.VB_VarHelpID = -1
Private WithEvents moHosts As CHosts
Attribute moHosts.VB_VarHelpID = -1

Private mlIntervalMinutes As Long
Private mlTickCount As Long
Private mbShouldLog As Boolean
Private msLogPath As String

Private Type HOST_INFO
    HostName As String
    LastIP As String
End Type

Event Enabled(ByVal isEnabled As Boolean)
Event StartIPChecking()
Event BeginIPLookUp(ByVal sHostName As String)
Event IPChanged(ByVal sKey As String, ByVal sHostName As String, ByVal sOLDIPaddress As String, ByVal sNewIPAddress As String)
Event AfterIPLookUp(ByVal sHostName As String, ByVal sIPAddress As String)
Event EndIPChecking()
Event ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
Event DataLoaded()
Event DataSaved()
Event HostAdded(ByVal sHostName As String)
Event HostRemoved(ByVal vIndex As Variant)
Event HostsCleared()
Public Property Let Enabled(ByVal RHS As Boolean)
    If moHosts.Count = 0 And RHS = True Then
        mErrorHandle 80001, "Enabled", "Can not enable object with no hosts"
        Exit Property
    End If
    
    If RHS = True Then
        mlTickCount = 0
        mCheckIPAddresses
    End If
    
    moTimer.Enabled = RHS
    
    RaiseEvent Enabled(RHS)
End Property
Public Property Get Enabled() As Boolean
    Enabled = moTimer.Enabled
End Property
Public Property Get Hosts() As CHosts
    Set Hosts = moHosts
End Property
Public Property Let IntervalMinutes(ByVal RHS As Long)
    If RHS < 1 Then Exit Property
    mlIntervalMinutes = RHS
End Property
Public Property Get IntervalMinutes() As Long
    IntervalMinutes = mlIntervalMinutes
End Property
Public Sub LoadHosts(ByVal sFilename As String)

On Error GoTo ErrorHandle

Dim iNextFree As Integer
Dim sHost As String
Dim vFields As Variant

    If Len(Dir(sFilename, vbNormal)) = 0 Then
        mErrorHandle 80002, "LoadHosts", "Could not find file"
        GoTo ClearVariables
    End If
    
    moHosts.Clear
    iNextFree = FreeFile
    Open sFilename For Input As iNextFree
        While Not EOF(iNextFree)
            Line Input #iNextFree, sHost
            vFields = Split(sHost, ",")
            moHosts.Add CStr(vFields(0)), CStr(vFields(1)), CDate(vFields(2)), CStr(vFields(3))
        Wend
    Close iNextFree

    RaiseEvent DataLoaded

ClearVariables:
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, "LoadHosts", Err.Description
    GoTo ClearVariables

End Sub
Public Property Let LogPath(ByVal RHS As String)
    If Len(RHS) = 0 Then Exit Property
    msLogPath = RHS
End Property
Public Property Get LogPath() As String
    LogPath = msLogPath
End Property
Private Sub mCheckIPAddresses()

On Error GoTo ErrorHandle

Dim oLogFile As New cLogFile
Dim oHostInfo As CHostInfo
Dim sIP As String
Dim sHost As String
Dim sOldIP As String
Dim bWasEnabled As Boolean

    oLogFile.FileName = msLogPath
    oLogFile.MaxSize = 3145728
    mlTickCount = 0
    
    RaiseEvent StartIPChecking
    If mbShouldLog = True Then oLogFile.AppendLog "Starting IP Checking"
    bWasEnabled = moTimer.Enabled
    moTimer.Enabled = False
    
    For Each oHostInfo In moHosts
        sIP = ""
        sHost = oHostInfo.HostName
        RaiseEvent BeginIPLookUp(sHost)
        If mbShouldLog = True Then oLogFile.AppendLog vbTab & "Looking up ip address for " & sHost
            sIP = moGetIPAddress.GetIPAddress(sHost)
            oHostInfo.LastChecked = Now
            
            If sIP <> oHostInfo.LastIPAddress Then
                sOldIP = oHostInfo.LastIPAddress
                oHostInfo.LastIPAddress = sIP
                RaiseEvent IPChanged(oHostInfo.Key, oHostInfo.HostName, sOldIP, sIP)
                If mbShouldLog = True Then oLogFile.AppendLog vbTab & vbTab & "!!!! Ip address for " & sHost & " has changed from " & sOldIP & " to " & sIP & " !!!"
                
            Else
                If mbShouldLog = True Then oLogFile.AppendLog vbTab & vbTab & "Ip address for " & sHost & " is " & sIP
            End If
            
        RaiseEvent AfterIPLookUp(sHost, sIP)
        
    Next

ClearVariables:
    RaiseEvent EndIPChecking
    If mbShouldLog = True Then oLogFile.AppendLog "Ending IP Checking"
    Set oLogFile = Nothing
    moTimer.Enabled = bWasEnabled
    
    Set oHostInfo = Nothing
    
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, "mCheckIPAddresses", Err.Description
    GoTo ClearVariables

End Sub
Private Sub mErrorHandle(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    RaiseEvent ErrorOccured(lErrorNumber, "CMonitorIPAddresses." & sErrorSource, sErrorDescription)
End Sub
Public Property Get MinutesLeft() As Long
    MinutesLeft = mlIntervalMinutes - mlTickCount
End Property
Public Sub SaveHosts(ByVal sFilename As String)

On Error GoTo ErrorHandle

Dim iNextFree As Integer
Dim oHostInfo As CHostInfo

    iNextFree = FreeFile

    Open sFilename For Output As iNextFree
        For Each oHostInfo In moHosts
            Print #iNextFree, oHostInfo.HostName & "," & oHostInfo.LastIPAddress & "," & oHostInfo.LastChecked & "," & oHostInfo.HOSTSEntry
        Next
    Close iNextFree

    RaiseEvent DataSaved
ClearVariables:
    Set oHostInfo = Nothing
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, "SaveHosts", Err.Description
    GoTo ClearVariables
    
End Sub
Public Property Let ShouldLog(ByVal RHS As Boolean)
    mbShouldLog = RHS
End Property
Public Property Get ShouldLog() As Boolean
    ShouldLog = mbShouldLog
End Property
Public Sub UpdateNow()
    mCheckIPAddresses
End Sub
Private Sub Class_Initialize()
    Set moHosts = New CHosts
    Set moTimer = New XTimer
    Set moGetIPAddress = New CGetIPAddress

    mlIntervalMinutes = 60
    mlTickCount = 0
    moTimer.Interval = 60000
    moTimer.Enabled = False
    
    msLogPath = App.Path & "\ip monitor log.txt"
End Sub
Private Sub Class_Terminate()
    Set moHosts = Nothing
    Set moGetIPAddress = Nothing
    moTimer.Enabled = False
    Set moTimer = Nothing
End Sub

Private Sub moGetIPAddress_ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    RaiseEvent ErrorOccured(lErrorNumber, sErrorSource, sErrorDescription)
End Sub

Private Sub moHosts_HostAdded(ByVal sKey As String)
    RaiseEvent HostAdded(sKey)
End Sub
Private Sub moHosts_HostRemoved(ByVal vIndex As Variant)
    RaiseEvent HostRemoved(vIndex)
End Sub

Private Sub moHosts_HostsCleared()
    RaiseEvent HostsCleared
End Sub
Private Sub moTimer_Tick()

    mlTickCount = mlTickCount + 1
    If mlTickCount >= mlIntervalMinutes Then
        mCheckIPAddresses
        mlTickCount = 0
    End If
    
End Sub
