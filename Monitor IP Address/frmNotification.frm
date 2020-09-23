VERSION 5.00
Begin VB.Form frmNotification 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Host has changed IP Address"
   ClientHeight    =   1050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmNotification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2400
      Top             =   240
   End
   Begin VB.CommandButton cmdUpdateHosts 
      Caption         =   "Update HOSTS file."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "At [DATE] [HOST] changed its IP address from [OLD IP] to [NEW IP]"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4395
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim msHOSTS As String
Dim msIP As String
Dim msSoundFile As String
Private Sub cmdUpdateHosts_Click()
    UpdateHOSTS msHOSTS, msIP
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
Public Sub DisplayChange(ByVal sOldIP As String, ByVal oHostInfo As CHostInfo, ByVal sSoundFile As String, ByVal bShouldPlay As Boolean, ByVal bShouldLoop As Boolean, ByVal oParent As Form)

    Me.lblInfo.Caption = Replace(Me.lblInfo.Caption, "[DATE]", oHostInfo.LastChecked)
    Me.lblInfo.Caption = Replace(Me.lblInfo.Caption, "[HOST]", oHostInfo.HostName)
    Me.lblInfo.Caption = Replace(Me.lblInfo.Caption, "[OLD IP]", sOldIP)
    Me.lblInfo.Caption = Replace(Me.lblInfo.Caption, "[NEW IP]", oHostInfo.LastIPAddress)
    
    msSoundFile = sSoundFile
    msHOSTS = oHostInfo.HOSTSEntry
    msIP = oHostInfo.LastIPAddress
    
    If bShouldPlay = True Then PlayNotifySound msSoundFile
    If bShouldLoop = True Then Me.Timer1.Enabled = True
    
    If Len(oHostInfo.HOSTSEntry) <> 0 Then Me.cmdUpdateHosts.Visible = True
    Me.Show , oParent
    
End Sub
Private Sub Timer1_Timer()
    PlayNotifySound msSoundFile
End Sub
