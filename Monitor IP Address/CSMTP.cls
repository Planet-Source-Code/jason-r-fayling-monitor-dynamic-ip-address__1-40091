VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents moWinsock As Winsock
Attribute moWinsock.VB_VarHelpID = -1
Private msBody As String
Private msSubject As String 'local copy
Private msReceipent As String 'local copy
Private msSender As String 'local copy
Private msSMTPHost As String 'local copy
Private moAttachedFiles As New Collection
Private mlState As SMTP_State
Private msEncodedFiles As String

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State

Event ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
Event ServerResponse(ByVal sMessage As String)
Event BeforeSend()
Event AfterSend(ByVal bSuccess As Boolean)
Public Property Get AttachedFiles() As Collection
    Set AttachedFiles = moAttachedFiles
End Property
Private Sub mLog(ByVal sMessage As String)
    RaiseEvent ServerResponse(sMessage)
End Sub
Public Function Send() As Boolean

On Error GoTo ErrorHandle

Dim vFile As Variant
Dim dEndTime As Double
Dim bSuccess As Boolean

    RaiseEvent BeforeSend
    
    For Each vFile In moAttachedFiles
        If Len(Dir(vFile, vbNormal)) <> 0 Then
            msEncodedFiles = msEncodedFiles & UUEncodeFile(CStr(vFile))
        End If
    Next
    
    moWinsock.Connect Trim(msSMTPHost), 25
    mlState = MAIL_CONNECT
    
    dEndTime = Timer + 90
    Do While mlState <> MAIL_QUIT And Timer < dEndTime
        DoEvents
    Loop
    
    If mlState = MAIL_QUIT Then bSuccess = True
    RaiseEvent AfterSend(bSuccess)
    
    Send = bSuccess
ClearVariables:
    Exit Function
    
ErrorHandle:
    mErrorHandle Err.Number, "Send", Err.Description
    GoTo ClearVariables

End Function
Private Sub mErrorHandle(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    RaiseEvent ErrorOccured(lErrorNumber, "CSMTP." & sErrorSource, sErrorDescription)
End Sub
Public Property Let SMTPHost(ByVal RHS As String)
    msSMTPHost = RHS
End Property
Public Property Get SMTPHost() As String
    SMTPHost = msSMTPHost
End Property
Public Property Let Sender(ByVal RHS As String)
    msSender = RHS
End Property
Public Property Get Sender() As String
    Sender = msSender
End Property
Public Property Let Receipent(ByVal RHS As String)
    msReceipent = RHS
End Property
Public Property Get Receipent() As String
    Receipent = msReceipent
End Property
Public Property Let subject(ByVal RHS As String)
    msSubject = RHS
End Property
Public Property Get subject() As String
    subject = msSubject
End Property
Public Property Set Winsock(ByRef RHS As Winsock)
    Set moWinsock = RHS
End Property

Private Sub Class_Terminate()
    moWinsock.Close
    Set moWinsock = Nothing
    Set moAttachedFiles = Nothing
End Sub
Public Property Get Body() As String
    Body = msBody
End Property
Public Property Let Body(ByVal RHS As String)
    msBody = RHS
End Property

Private Function UUEncodeFile(strFilePath As String) As String

    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    Dim strResult       As String
    '
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    '
    'Insert first marker: "begin 664 ..."
    strResult = "begin 664 " + strFileName + vbLf
    '
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize \ 45 + 1
    '
    'Prepare buffer to retrieve data from
    'the file by 45 symbols chunks
    strFileData = Space(45)
    '
    intFile = FreeFile
    '
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            'Read file data by 45-bytes cnunks
            '
            If i = lEncodedLines Then
                'Last line of encoded data often is not
                'equal to 45, therefore we need to change
                'size of the buffer
                strFileData = Space(lFileSize Mod 45)
            End If
            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            '
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                'If the last line is processed and length of
                'source data is not a number divisible by 3, add one or two
                'blankspace symbols
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            strResult = strResult + strTempLine + vbLf
            'reset line buffer
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    strResult = strResult & "`" & vbLf + "end" + vbLf
    'asign return value
    UUEncodeFile = strResult
    
End Function

Private Sub moWinsock_DataArrival(ByVal bytesTotal As Long)

Dim sServerResponse   As String
Dim sResponseCode     As String
Dim sDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    moWinsock.GetData sServerResponse
    '
    mLog sServerResponse
    '
    'Get server response code (first three symbols)
    '
    sResponseCode = Left(sServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If sResponseCode = "250" Or _
       sResponseCode = "220" Or _
       sResponseCode = "354" Then
       
        Select Case mlState
            Case MAIL_CONNECT
                'Change current state of the session
                mlState = MAIL_HELO
                '
                'Remove blank spaces
                sDataToSend = Trim$(msSender)
                '
                'Retrieve mailbox name from e-mail address
                sDataToSend = Left$(sDataToSend, _
                                InStr(1, sDataToSend, "@") - 1)
                'Send HELO command to the server
                moWinsock.SendData "HELO " & sDataToSend & vbCrLf
                '
                mLog "HELO " & sDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                mlState = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                moWinsock.SendData "MAIL FROM:" & Trim$(msSender) & vbCrLf
                '
                mLog "MAIL FROM:" & Trim$(msSender)
                '
            Case MAIL_FROM
                '
                '
                'Send RCPT TO command to the server
                Dim vFields As Variant
                Static lCurrent As Long
                Dim lTotal As Long
                
                vFields = Split(Replace(msReceipent, ";", ","), ",")
                lTotal = UBound(vFields)
                
                moWinsock.SendData "RCPT TO:" & vFields(lCurrent) & vbCrLf
                lCurrent = lCurrent + 1
                
                If lCurrent > lTotal Then
                    'Change current state of the session
                    mlState = MAIL_RCPTTO
                End If
                
                '
                mLog "RCPT TO:" & Trim$(msReceipent)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                mlState = MAIL_DATA
                '
                'Send DATA command to the server
                moWinsock.SendData "DATA" & vbCrLf
                '
                mLog "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                mlState = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf
                '
                'Send Subject line
                moWinsock.SendData "Subject:" & msSubject & vbLf & vbCrLf
                '
                mLog "Subject:" & msSubject
                '
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                '
                'Add atacchments
                strMessage = msBody & vbCrLf & vbCrLf & msEncodedFiles
                'clear memory
                msEncodedFiles = ""
                'Parse message to get lines (for VB6 only)
                varLines = Split(strMessage, vbCrLf)
                'clear memory
                strMessage = ""
                '
                'Send each line of the message
                For Each varLine In varLines
                    moWinsock.SendData CStr(varLine) & vbLf
                    '
                    mLog CStr(varLine)
                Next
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                moWinsock.SendData vbCrLf & "." & vbCrLf
                '
                mLog "."
                '
            Case MAIL_DOT
                'Change current state of the session
                mlState = MAIL_QUIT
                '
                'Send QUIT command to the server
                moWinsock.SendData "QUIT" & vbCrLf
                '
                mLog "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                moWinsock.Close
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        moWinsock.Close
        '
        If Not mlState = MAIL_QUIT Then
            mLog "SMTP Error: " & sServerResponse
        Else
            mLog "Message sent successfuly."
        End If
        '
    End If

End Sub
Private Sub moWinsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    mErrorHandle Number, "moWinsock_Error", Description
End Sub


