VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHostInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msHostName As String
Private msLastIPAddress As String
Private msLastChecked As Date
Private msHOSTSEntry As String
Public Property Get HostName() As String
    HostName = msHostName
End Property
Friend Property Let HostName(ByVal RHS As String)
    msHostName = RHS
End Property
Public Property Get LastIPAddress() As String
    LastIPAddress = msLastIPAddress
End Property
Friend Property Let LastIPAddress(ByVal RHS As String)
    msLastIPAddress = RHS
End Property
Public Property Get Key() As String
    Key = "KEY_" & msHostName
End Property
Public Property Get LastChecked() As Date
    LastChecked = msLastChecked
End Property
Friend Property Let LastChecked(ByVal RHS As Date)
    msLastChecked = RHS
End Property
Public Property Get HOSTSEntry() As String
    HOSTSEntry = msHOSTSEntry
End Property
Friend Property Let HOSTSEntry(ByVal RHS As String)
    msHOSTSEntry = RHS
End Property
