VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitorIPAddress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor IP Address"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmMonitorIPAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdateNow 
      Caption         =   "&Update now"
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   4635
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8890
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRemoveHost 
      Caption         =   "&Remove Host"
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddHost 
      Caption         =   "&Add Host"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.ListView listHosts 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Host Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Checked"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "HOSTS"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Begin Monitoring"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtQuery 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "60"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Frame frame1 
      Caption         =   "Actions"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4815
      Begin VB.CheckBox chkUpdateHOSTS 
         Caption         =   "Update HOSTS entry."
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkLoopSound 
         Caption         =   "Loop Sound"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfigureEmail 
         Caption         =   "Configure"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkSendEmail 
         Caption         =   "Send an email."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdLookForSound 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySound 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkPlaySound 
         Caption         =   "Play a sound"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkDisplayNotification 
         Caption         =   "Display a notification window"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check every "
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "minutes."
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   4200
      Width           =   585
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hosts:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpStartMonitor 
         Caption         =   "Begin Monitoring"
      End
      Begin VB.Menu mnuPopUpUpdateNow 
         Caption         =   "Update Now"
      End
      Begin VB.Menu mnuPopUpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuPopUpExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMonitorIPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moMonitorIPAddresses As CMonitorIPAddresses
Attribute moMonitorIPAddresses.VB_VarHelpID = -1
Private WithEvents moSMTP As CSMTP
Attribute moSMTP.VB_VarHelpID = -1
Private WithEvents moSysTray As SysTray
Attribute moSysTray.VB_VarHelpID = -1

Private msSoundFile As String
Private msSMTP As String
Private msSender As String
Private msSendTo As String

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Sub GetSettings()

Dim sFile As String
Dim sSection As String
Dim sData As String
Dim iLength As Long
    
    sFile = App.Path & "\settings.dat"
    sSection = "MONITOR IP ADDRESS"
    
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "displayNotifyWindow", "1", sData, Len(sData), sFile)
        Me.chkDisplayNotification.Value = Val(Left(sData, iLength))
    
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "playSound", "0", sData, Len(sData), sFile)
        Me.chkPlaySound.Value = Val(Left(sData, iLength))
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "sound", "", sData, Len(sData), sFile)
        msSoundFile = Left(sData, iLength)
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "loopSound", "0", sData, Len(sData), sFile)
        Me.chkLoopSound.Value = Val(Left(sData, iLength))

    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "updateHosts", "0", sData, Len(sData), sFile)
        Me.chkUpdateHOSTS.Value = Val(Left(sData, iLength))
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "queryIntervals", "60", sData, Len(sData), sFile)
        Me.txtQuery.Text = Val(Left(sData, iLength))
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "sendEmail", "0", sData, Len(sData), sFile)
        Me.chkSendEmail.Value = Val(Left(sData, iLength))
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "smtp", "", sData, Len(sData), sFile)
        msSMTP = Left(sData, iLength)
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "sender", "", sData, Len(sData), sFile)
        msSender = Left(sData, iLength)
        
    sData = String(255, " ")
    iLength = GetPrivateProfileString(sSection, "sendTo", "", sData, Len(sData), sFile)
        msSendTo = Left(sData, iLength)

End Sub
Private Sub mEnableControls(ByVal bEnable As Boolean)

On Error GoTo ErrorHandle:
    
    Me.cmdAddHost.Enabled = bEnable
    Me.cmdRemoveHost.Enabled = bEnable
    Me.txtQuery.Enabled = bEnable
    Me.frame1.Enabled = bEnable
    
    If bEnable = True And Me.listHosts.ListItems.Count = 0 Then Me.cmdRemoveHost.Enabled = False

ClearVariables:
    Exit Sub
    
ErrorHandle:
    GoTo ClearVariables

End Sub
Private Sub mErrorHandle(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    MsgBox sErrorDescription & vbCrLf & vbCrLf & sErrorSource, vbInformation, CStr("Error Number: " & lErrorNumber)
End Sub
Private Sub mFillGrid()

On Error GoTo ErrorHandle

Dim oHostInfo As CHostInfo
Dim oListItem As ListItem

    Me.listHosts.ListItems.Clear
    
    For Each oHostInfo In moMonitorIPAddresses.Hosts
        Set oListItem = Me.listHosts.ListItems.Add(, oHostInfo.Key, oHostInfo.HostName)
            oListItem.ListSubItems.Add , , oHostInfo.LastIPAddress
            oListItem.ListSubItems.Add , , oHostInfo.LastChecked
            oListItem.ListSubItems.Add , , oHostInfo.HOSTSEntry
    Next
    
    If Me.listHosts.ListItems.Count > 0 And moMonitorIPAddresses.Enabled = False Then
        Me.cmdStartStop.Enabled = True
        Me.cmdRemoveHost.Enabled = True
    Else
        Me.cmdStartStop.Enabled = False
        Me.cmdRemoveHost.Enabled = False
    End If
    
ClearVariables:
    AutoSizeColumns Me.listHosts
    Set oListItem = Nothing
    Set oHostInfo = Nothing
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, "frmMonitorIPAddress.mfillgrid", Err.Description
    GoTo ClearVariables
    
End Sub

Private Sub SaveSettings()
    
Dim sFile As String
Dim sSection As String
    
    sFile = App.Path & "\settings.dat"
    sSection = "MONITOR IP ADDRESS"
    
    WritePrivateProfileString sSection, "displayNotifyWindow", CStr(Me.chkDisplayNotification.Value), sFile
    WritePrivateProfileString sSection, "playSound", CStr(Me.chkPlaySound.Value), sFile
    WritePrivateProfileString sSection, "sound", msSoundFile, sFile
    WritePrivateProfileString sSection, "loopSound", CStr(Me.chkLoopSound.Value), sFile
    WritePrivateProfileString sSection, "updateHOSTS", CStr(Me.chkUpdateHOSTS.Value), sFile
    WritePrivateProfileString sSection, "queryIntervals", CStr(Val(Me.txtQuery.Text)), sFile
    WritePrivateProfileString sSection, "sendEmail", CStr(Me.chkSendEmail.Value), sFile
    WritePrivateProfileString sSection, "smtp", msSMTP, sFile
    WritePrivateProfileString sSection, "sender", msSender, sFile
    WritePrivateProfileString sSection, "sendTo", msSendTo, sFile
    
End Sub

Private Sub chkDisplayNotification_Click()
    
    Me.chkPlaySound.Enabled = Me.chkDisplayNotification.Value
    If Me.chkDisplayNotification.Value = 0 Then
        Me.chkPlaySound.Value = 0
        Me.chkLoopSound.Value = 0
    End If
    
End Sub

Private Sub chkPlaySound_Click()
    Me.cmdPlaySound.Enabled = Me.chkPlaySound.Value
    Me.cmdLookForSound.Enabled = Me.chkPlaySound.Value
    Me.chkLoopSound.Enabled = Me.chkPlaySound.Value
End Sub

Private Sub chkSendEmail_Click()
    Me.cmdConfigureEmail.Enabled = Me.chkSendEmail.Value
End Sub
Private Sub cmdAddHost_Click()

On Error GoTo ErrorHandle

Dim sHostName As String
Dim sHOSTSENTRY As String

    sHostName = InputBox("Host Name", "What is the name of the server you wish to monitor?", "")
    If Len(sHostName) = 0 Then GoTo ClearVariables
    
    If MsgBox("Do you wish to associate an entry in your HOSTS file with this server?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        sHOSTSENTRY = InputBox("HOSTS Entry", "What is the HOSTS name?", sHostName)
    End If
    
    moMonitorIPAddresses.Hosts.Add sHostName, , , sHOSTSENTRY

ClearVariables:
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, "frmMonitorIPAddress.cmdAddHost_Click", Err.Description
    GoTo ClearVariables
End Sub

Private Sub cmdConfigureEmail_Click()

    frmSMTP.ConfigureSMTP msSMTP, msSender, msSendTo, Me

End Sub

Private Sub cmdLookForSound_Click()

    With Me.CommonDialog1
        .CancelError = False
        .DefaultExt = ".wav"
        .DialogTitle = "Wave File"
        .FileName = msSoundFile
        .Filter = "Wav Files (*.wav)|*.wav"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        msSoundFile = .FileName
    End With

End Sub

Private Sub cmdPlaySound_Click()
    PlayNotifySound msSoundFile
End Sub

Private Sub cmdRemoveHost_Click()

On Error GoTo ErrorHandle

Dim sKey As String
Dim sHost As String

    If Me.listHosts.SelectedItem Is Nothing Then GoTo ClearVariables

    sKey = Me.listHosts.SelectedItem.Key
    sHost = Me.listHosts.SelectedItem.Text
    
    If MsgBox("Are you sure you wish to remove the host: " & sHost & " from your list?", vbQuestion + vbYesNo + vbDefaultButton2, "Remove Host?") = vbNo Then GoTo ClearVariables
    moMonitorIPAddresses.Hosts.Remove sKey

ClearVariables:
    Exit Sub
    
ErrorHandle:
    mErrorHandle Err.Number, Err.Source, Err.Description
    GoTo ClearVariables

End Sub
Private Sub cmdStartStop_Click()
    
    If Val(Me.cmdStartStop.Tag) = 0 Then
        Me.cmdStartStop.Tag = 1
        Me.cmdStartStop.Caption = "Stop Monitoring"
        With moMonitorIPAddresses
            .IntervalMinutes = Val(Me.txtQuery.Text)
            .Enabled = True
        End With
        mEnableControls False
    Else
        Me.cmdStartStop.Tag = 0
        Me.cmdStartStop.Caption = "Begin Monitoring"
        With moMonitorIPAddresses
            .Enabled = False
        End With
        mEnableControls True
        Me.StatusBar1.Panels(1).Text = ""
    End If
    
    Me.mnuPopUpStartMonitor.Caption = Me.cmdStartStop.Caption
    
End Sub
Private Sub cmdUpdateNow_Click()
    moMonitorIPAddresses.UpdateNow
End Sub

Private Sub Form_Load()
    
    Set moMonitorIPAddresses = New CMonitorIPAddresses
        With moMonitorIPAddresses
            .LoadHosts App.Path & "\hosts.dat"
            .ShouldLog = True
            .LogPath = App.Path & "\iplog.txt"
        End With
        
    
    Set moSMTP = New CSMTP
        Set moSMTP.Winsock = Me.Winsock1
    
    Set moSysTray = New SysTray
        With moSysTray
            .Form = Me
            .Icon = Me.Icon
            .Persistent = True
            .PopupMenu = Me.mnuPopUp
            .PopupStyle = stOnRightDown
            .RestoreFromTrayOn = stOnLeftDblClick Or stOnRightDblClick
            .TrayFormStyle = stHideFormWhenMin
            .TrayTip = Me.Caption
            .Visible = True
        End With
        
    GetSettings
    mFillGrid
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    moSysTray.UnloadTray
    Set moSysTray = Nothing
    
    moMonitorIPAddresses.SaveHosts App.Path & "\hosts.dat"
    SaveSettings
    Set moMonitorIPAddresses = Nothing
    
End Sub

Private Sub mnuPopUpExit_Click()
    Unload Me
End Sub

Private Sub mnuPopUpRestore_Click()
    moSysTray.FormRestore
End Sub

Private Sub mnuPopUpStartMonitor_Click()
    cmdStartStop_Click
End Sub
Private Sub mnuPopUpUpdateNow_Click()
    moMonitorIPAddresses.UpdateNow
End Sub
Private Sub moMonitorIPAddresses_DataLoaded()
    mFillGrid
End Sub
Private Sub moMonitorIPAddresses_Enabled(ByVal isEnabled As Boolean)
    If isEnabled = True Then
        Me.StatusBar1.Panels(1).Text = "Next update will be at " & Now + (0.00001 * (60 * moMonitorIPAddresses.MinutesLeft))
    Else
        Me.StatusBar1.Panels(1).Text = ""
    End If
End Sub
Private Sub moMonitorIPAddresses_EndIPChecking()
    mFillGrid
End Sub
Private Sub moMonitorIPAddresses_ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    If lErrorNumber = 80002 Then Exit Sub
    mErrorHandle lErrorNumber, sErrorSource, sErrorDescription
End Sub

Private Sub moMonitorIPAddresses_HostAdded(ByVal sHostName As String)
    mFillGrid
End Sub

Private Sub moMonitorIPAddresses_HostRemoved(ByVal vIndex As Variant)
    mFillGrid
End Sub
Private Sub moMonitorIPAddresses_HostsCleared()
    mFillGrid
End Sub
Private Sub moMonitorIPAddresses_IPChanged(ByVal sKey As String, ByVal sHostName As String, ByVal sOLDIPaddress As String, ByVal sNewIPAddress As String)

Dim oForm As frmNotification
Dim oHostInfo As CHostInfo
    
    If sOLDIPaddress = "0.0.0.0" Then Exit Sub

    Set oHostInfo = moMonitorIPAddresses.Hosts(sKey)
    
    If Me.chkDisplayNotification.Value = 1 Then
        Set oForm = New frmNotification
        oForm.DisplayChange sOLDIPaddress, oHostInfo, msSoundFile, Me.chkPlaySound.Value, Me.chkLoopSound.Value, Me
    End If
    
    If Me.chkUpdateHOSTS.Value = 1 Then
        UpdateHOSTS oHostInfo.HOSTSEntry, sNewIPAddress
    End If
    
    If Me.chkSendEmail.Value = 1 Then
        With moSMTP
            .Body = "At " & oHostInfo.LastChecked & " " & oHostInfo.HostName & " changed its IP address from " & sOLDIPaddress & " to " & sNewIPAddress
            .subject = "IP Address Changed"
            .Receipent = msSendTo
            .Sender = msSender
            .SMTPHost = msSMTP
            .Send
        End With
    End If
    
    Set oHostInfo = Nothing

End Sub
Private Sub moMonitorIPAddresses_StartIPChecking()
    If moMonitorIPAddresses.Enabled = True Then
        Me.StatusBar1.Panels(1).Text = "Next update will be at " & Now + (0.00001 * (60 * moMonitorIPAddresses.MinutesLeft))
    End If
End Sub

Private Sub moSMTP_ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    mErrorHandle lErrorNumber, sErrorSource, sErrorDescription
End Sub
Private Sub txtQuery_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub
