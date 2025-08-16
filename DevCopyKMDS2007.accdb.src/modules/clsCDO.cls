VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Const cdoSendUsingPickup = 1    'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2      'Send the message using the network (SMTP over the network).

Const cdoNotUsed = -1

Public Enum AuthTypes
    cdoAnonymous = 0  'Do not authenticate
    cdoBasic = 1      'Clear-text authentication
    cdoNTLM = 2       'NT authentication
End Enum


Public Server       As String           'Name or IP of Remote SMTP Server
Public UserId       As String           'Logon name
Public Password     As String           'Logon password
Public SendUsing    As Long             'default: cdoSendUsingPort
Public Authenticate As AuthTypes        'default: cdoNotUsed
Public ServerPort   As Long             'default: cdoNotUsed
Public UseSSL       As Long             'default: cdoNotUsed
Public Timeout      As Long             'default: cdoNotUsed
'
Public Sender       As String           'eg billyg@microsoft.com
Public SenderName   As String           'eg Bill Gates
Public Receiver     As String           'eg smccall@myaddress.com
Public ReceiverName As String           'eg Stuart McCall
Public CC           As String           'Carbon Copy
Public BCC          As String           'Blind Carbon Copy
Public Subject      As String           'Message subject text
Public Text         As String           'Message body text
Public HTML         As Boolean          'Indicates message body text is HTML
Public Files        As New Collection   'File paths of files to attach

Public Sub SendMail()
    Const qt = """"
    Const Prefix = "http://schemas.microsoft.com/cdo/configuration/"
    Dim msg As Object
    Dim f As Variant

    Set msg = CreateObject("CDO.Message")
    With msg
        If Not IsVoid(SenderName) Then
            .From = qt & SenderName & qt & " <" & Sender & ">"
        Else
            .From = Sender
        End If
        If Not IsVoid(ReceiverName) Then
            .To = qt & ReceiverName & qt & " <" & Receiver & ">"
        Else
            .To = Receiver
        End If
        .CC = CC
        .BCC = BCC
        .Subject = Subject
        If HTML Then .HTMLBody = Text Else .TextBody = Text
        For Each f In Files
            .AddAttachment f
        Next
        With .Configuration.Fields
            .Item(Prefix & "sendusing") = SendUsing
            If Not IsVoid(Server) Then .Item(Prefix & "smtpserver") = Server
            If Not IsVoid(UserId) Then .Item(Prefix & "sendusername") = UserId
            If Not IsVoid(Password) Then .Item(Prefix & "sendpassword") = Password
            If Authenticate > cdoNotUsed Then .Item(Prefix & "smtpauthenticate") = Authenticate
            If ServerPort > cdoNotUsed Then .Item(Prefix & "smtpserverport") = ServerPort
            If UseSSL > cdoNotUsed Then .Item(Prefix & "smtpusessl") = UseSSL
            If Timeout > cdoNotUsed Then .Item(Prefix & "smtpconnectiontimeout") = Timeout
            .Update
        End With
        .Send
    End With

    Set msg = Nothing
End Sub

Private Sub Class_Initialize()
    SendUsing = cdoSendUsingPort
    Authenticate = cdoNotUsed
    ServerPort = cdoNotUsed
    UseSSL = cdoNotUsed
    Timeout = cdoNotUsed
End Sub

Private Function IsVoid(v) As Boolean
    IsVoid = (Len(v & "") = 0)
End Function
