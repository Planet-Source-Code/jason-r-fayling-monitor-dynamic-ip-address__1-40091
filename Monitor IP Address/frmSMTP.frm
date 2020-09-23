VERSION 5.00
Begin VB.Form frmSMTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMTP Setup"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5505
   Icon            =   "frmSMTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTo 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtSMTP 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the email addresses you wish to send the notification to below, seperate with a semicolan."
      Height          =   555
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email Address:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Host:"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private msSMTP As String
Private msSender As String
Private msTo As String
Private mbCancelled As Boolean
Private Sub CancelButton_Click()
    mbCancelled = True
    Unload Me
End Sub
Public Sub ConfigureSMTP(ByRef SMTPHost As String, ByRef Sender As String, ByRef SendTo As String, ByVal oParent As Form)

    msSMTP = SMTPHost
    msSender = Sender
    msTo = SendTo
    
    Me.txtSMTP.Text = msSMTP
    Me.txtSender.Text = msSender
    Me.txtTo.Text = msTo
    
    Me.Show 1, oParent

    If mbCancelled = False Then
        SMTPHost = msSMTP
        Sender = msSender
        SendTo = msTo
    End If
    
    Unload Me

End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub


Private Sub txtSender_Change()
    msSender = Me.txtSender.Text
End Sub
Private Sub txtSMTP_Change()
    msSMTP = Me.txtSMTP.Text
End Sub

Private Sub txtTo_Change()
    msTo = Me.txtTo.Text
End Sub
