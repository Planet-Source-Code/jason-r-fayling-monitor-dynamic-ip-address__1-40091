VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CHostInfo"
Attribute VB_Ext_KEY = "Member0" ,"CHostInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private moHosts As Collection

Event HostAdded(ByVal sKey As String)
Event HostRemoved(ByVal vIndex As Variant)
Event HostsCleared()

Public Function Add(HostName As String, Optional ByVal LastIPAddress As String, Optional ByVal LastChecked As Date = "01/01/1900", Optional ByVal HOSTSEntry As String) As CHostInfo

Dim oHostInfo As CHostInfo
    
    Set oHostInfo = New CHostInfo

    If Len(LastIPAddress) = 0 Then LastIPAddress = "0.0.0.0"
    
    oHostInfo.LastIPAddress = LastIPAddress
    oHostInfo.HostName = HostName
    oHostInfo.LastChecked = LastChecked
    oHostInfo.HOSTSEntry = HOSTSEntry
    
    moHosts.Add oHostInfo, "KEY_" & HostName
    RaiseEvent HostAdded("KEY_" & HostName)
    'return the object created
    Set Add = oHostInfo
    Set oHostInfo = Nothing

End Function
Public Property Get Item(vntIndexKey As Variant) As CHostInfo
Attribute Item.VB_UserMemId = 0
  Set Item = moHosts(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = moHosts.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    moHosts.Remove vntIndexKey
    RaiseEvent HostRemoved(vntIndexKey)
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = moHosts.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set moHosts = New Collection
End Sub
Private Sub Class_Terminate()
    Set moHosts = Nothing
End Sub
Public Sub Clear()
    Set moHosts = Nothing
    Set moHosts = New Collection
    RaiseEvent HostsCleared
End Sub
