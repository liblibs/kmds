Attribute VB_Name = "modEMail"
Option Compare Database
Option Explicit

'info from http://www.paulsadowski.com/WSH/cdo.htm
'
'Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
'Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
'
'Const cdoAnonymous = 0 'Do not authenticate
'Const cdoBasic = 1 'basic (clear-text) authentication
'Const cdoNTLM = 2 'NTLM
'
'Set CDOmsg = CreateObject("CDO.Message")
'CDOmsg.Subject = "Example CDO Message"
'CDOmsg.From = """Me"" <me@my.com>"
'CDOmsg.To = "test@paulsadowski.com"
'CDOmsg.TextBody = "This is some sample message text.." & vbcrlf & "It was sent using SMTP authentication."
'
''==This section provides the configuration information for the remote SMTP server.
'
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'
''Name or IP of Remote SMTP Server
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.your.com"
'
''Type of authentication, NONE, Basic (Base64 encoded), NTLM
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'
''Your UserID on the SMTP server
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendusername") = "youruserid"
'
''Your password on the SMTP server
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "yourpassword"
'
''Server port (typically 25)
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'
''Use SSL for the connection (False or True)
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
'
''Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
'CDOmsg.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'
'CDOmsg.Configuration.Fields.Update
'
''==End remote SMTP server configuration section==
'
'CDOmsg.Send



Public Function EmailFileToKMDS(FileName As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).

Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

Dim CDOmsg As New CDO.message

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso As New FileSystemObject
Dim f As TextStream
Dim bodytext As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If gstrStoreName = "" Then LoadDefaults


    'Open the file for reading
    Set f = fso.OpenTextFile(FileName, ForReading)
    'The ReadAll method reads the entire file into the variable BodyText
    bodytext = f.ReadAll
    'Close the file
    f.Close
    Set f = Nothing
    Set fso = Nothing


    With CDOmsg
        .From = GetKMDSSettings("KMDSMail from", "Agents")
        .To = "info@kmds.us"
        .Subject = "EODLog from " & gstrStoreName
        .TextBody = bodytext

        '==This section provides the configuration information for the remote SMTP server.

        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = GetKMDSSettings("KMDSMail sendusing", "2")

        'Name or IP of Remote SMTP Server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetKMDSSettings("KMDSMail smtpserver", "kmds.ipower.com")

        'Type of authentication, NONE, Basic (Base64 encoded), NTLM
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = GetKMDSSettings("KMDSMail smtpauthenticate", "1")

        'Your UserID on the SMTP server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GetKMDSSettings("KMDSMail sendusername", "Agents@kmds.us")

        'Your password on the SMTP server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GetKMDSSettings("KMDSMail sendpassword", "Ag3nt%xxxx")

        'Server port (typically 25)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GetKMDSSettings("KMDSMail smtpserverport", "587") '25

        'Use SSL for the connection (False or True)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = GetKMDSSettings("KMDSMail smtpusessl", "0")

        'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = GetKMDSSettings("KMDSMail smtpconnectiontimeout", "60")

        .Configuration.Fields.Update

        '==End remote SMTP server configuration section==

        .Send
    End With


ExitFunc:
    Exit Function

ErrLog:
    MsgBox Err.Number & " " & Err.Description & vbCrLf & FileName, , "EmailFileToKMDS"
    Resume ExitFunc

End Function



Public Sub KMDSeMail(mailto As String, FilesToAttach As Collection, Subject As String, Optional TextBodyOption As String)  'QueryName As String)

Dim strEmail As String
Dim curTotal As Currency
Dim CDOmsg As New CDO.message
Dim strTotal As String
Dim strStoreName As String
Dim strTextMSG As String
Dim bTextMsg As Boolean
Dim fsoTemp As New FileSystemObject
Dim varReturn As Variant
Dim FileName As String

Dim nI As Integer

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).

Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

    Screen.MousePointer = 11
    If gcfHandleErrors Then On Error GoTo ErrLog


    For Each varReturn In FilesToAttach
        If Not fsoTemp.FileExists(varReturn) Then
            MsgBox varReturn & vbCrLf & "File name does not exist", , "KMDS eMail"
            Exit Sub
        End If
        FileName = varReturn
    Next

    If InStr(mailto, "@") = 0 Then
        MsgBox mailto & " invalid EMail... not defined", vbCritical, "KMDSeMail"
        Exit Sub
    End If

    If gcfHandleErrors Then On Error GoTo ErrLogWithName


    With CDOmsg

        .From = GetKMDSSettings("KMDSMail from", "Agents@kmds.us")
        .To = mailto
        .Subject = Subject
        nI = 0
        For Each varReturn In FilesToAttach
            nI = nI + 1
            .AddAttachment varReturn
        Next
        If nI > 0 Then .TextBody = "Please find attached the information from " & gstrStoreName & vbCrLf
        If TextBodyOption <> "" Then
            .TextBody = .TextBody & TextBodyOption
        End If

        '==This section provides the configuration information for the remote SMTP server.

        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = GetKMDSSettings("KMDSMail sendusing", "2")

        'Name or IP of Remote SMTP Server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetKMDSSettings("KMDSMail smtpserver", "kmds.ipower.com")

        'Type of authentication, NONE, Basic (Base64 encoded), NTLM
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = GetKMDSSettings("KMDSMail smtpauthenticate", "1")

        'Your UserID on the SMTP server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GetKMDSSettings("KMDSMail sendusername", "Agents@kmds.us")

        'Your password on the SMTP server
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GetKMDSSettings("KMDSMail sendpassword", "Ag3nt$xxxx")

        'Server port (typically 25)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GetKMDSSettings("KMDSMail smtpserverport", "587") '25

        'Use SSL for the connection (False or True)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = GetKMDSSettings("KMDSMail smtpusessl", "0")

        'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = GetKMDSSettings("KMDSMail smtpconnectiontimeout", "60")

        .Configuration.Fields.Update

        '==End remote SMTP server configuration section==


    If gcfHandleErrors Then On Error GoTo ErrLog
        varReturn = SysCmd(acSysCmdSetStatus, "Emailing " & mailto & " " & FileName)
        .Send
    End With
    Set CDOmsg = Nothing

'    DoCmd.SendObject acSendNoObject, , , strEmail, , , "Store Sales as of " & Now, strSQL, False
ExitSub:
    Screen.MousePointer = 0
    SysCmd acSysCmdClearStatus
    EOLog "EODeMail Recap" & vbCrLf & CDOmsg.To & vbCrLf & CDOmsg.Subject & vbCrLf & CDOmsg.TextBody
Exit Sub

ErrLogWithName:

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description & " : KMDSeMail"
    Resume ExitSub
End Sub


Public Sub EODTextMessage() 'QueryName As String)
    'special e-mail designed primarily for text messages

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    Dim strSQL As String
    Dim strEmail As String
    Dim curTotal As Currency
    Dim CDOmsg As New CDO.message
    Dim strTotal As String
    Dim strStoreName As String
    Dim strTextMSG As String
    Dim bTextMsg As Boolean
    Dim fsoTemp As New FileSystemObject

    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).

    Const cdoAnonymous = 0 'Do not authenticate
    Const cdoBasic = 1 'basic (clear-text) authentication
    Const cdoNTLM = 2 'NTLM


    If gcfHandleErrors Then On Error GoTo ErrLog

    strSQL = "qryEOSalesForEMail"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = strSQL
    End With

    Set rstemp = comQDF.Execute
    If rstemp.EOF Then
        EOLog "No data to report"
        Exit Sub
    End If

    strEmail = rstemp!EMail
    If InStr(strEmail, "@") = 0 Then
        EOLog "Agent EMail is not defined"
        Exit Sub
    End If


    'format output for textmsg (120 max characters)
    If IsNumeric(Left(strEmail, 10)) Then
        bTextMsg = True
    Else
        bTextMsg = False
    End If

    strSQL = ""

    If gcfHandleErrors Then On Error GoTo ErrLogWithName
    If bTextMsg Then 'format for a text message
        With rstemp
            strStoreName = !AgentNumber & " " & Format(Now, "mm-dd hh:nn") & vbCrLf
            strSQL = strSQL & Format(" ", "!@@@@@@@@@@@@") & vbTab & "============" & vbCrLf
            Do While Not .EOF
                curTotal = curTotal + CCur(!TotalP)
                strSQL = strSQL & Format(!GroupDescription, "!@@@@@@@@@@@@") & vbTab & Format(Format(!TotalP, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
                If InStr(!GroupDescription, "LIQUOR") Then
                    strTextMSG = strTextMSG & Format(!GroupDescription, "!@@@@@@@@@@@@") & vbTab & Format(Format(!TotalP, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
                End If

                .MoveNext
            Loop
            strTotal = Format("== TOTAL ==", "!@@@@@@@@@@@@") & vbTab & Format(Format(curTotal, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
            .Close
            CDOmsg.Subject = ""
            CDOmsg.TextBody = Left(strStoreName & " " & strTotal & strTextMSG, 120)
        End With
    Else  'format for a email
        With rstemp
            strStoreName = !AgentNumber & " " & !Name & vbCrLf
            Do While Not .EOF
                curTotal = curTotal + CCur(!TotalP)
                strSQL = strSQL & Format(!GroupDescription, "!@@@@@@@@@@@@@@@@@@@@@@@@") & vbTab & Format(Format(!TotalP, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
                If InStr(!GroupDescription, "LIQUOR") Then
                    strTextMSG = strTextMSG & Format(!GroupDescription, "!@@@@@@@@@@@@@@@@@@@@@@@@") & vbTab & Format(Format(!TotalP, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
                End If
                .MoveNext
            Loop
            strSQL = strSQL & Format(" ", "!@@@@@@@@@@@@@@@@@@@@@@@@") & vbTab & "============" & vbCrLf
            strSQL = strSQL & Format("== TOTAL ==", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@") & vbTab & Format(Format(curTotal, "###,###.00"), "@@@@@@@@@@@@") & vbCrLf
            .Close
            CDOmsg.Subject = strStoreName & " as of " & Now
            CDOmsg.TextBody = strSQL

            'added 2009.05.07
'            If WhatIsMyIp Then
'                If fsoTemp.FileExists("c:\kmds\kmdsgetweb.txt") Then
'                    CDOmsg.AddAttachment "c:\kmds\kmdsgetweb.txt"
'                End If
'            End If

        End With
    End If

    Set rstemp = Nothing

    CDOmsg.From = GetKMDSSettings("KMDSMail from", "Agents@kmds.us")
    CDOmsg.To = strEmail
    With CDOmsg.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = GetKMDSSettings("KMDSMail sendusing", "2")
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetKMDSSettings("KMDSMail smtpserver", "kmds.ipower.com")
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = GetKMDSSettings("KMDSMail smtpauthenticate", "1")
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GetKMDSSettings("KMDSMail sendusername", "Agents@kmds.us")
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GetKMDSSettings("KMDSMail sendpassword", "Ag3nt$xxxx")
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = GetKMDSSettings("KMDSMail smtpserverport", "587") '25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = GetKMDSSettings("KMDSMail smtpusessl", "0")
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = GetKMDSSettings("KMDSMail smtpconnectiontimeout", "60")
        .Update
    End With
        '==End remote SMTP server configuration section==

    CDOmsg.Send

'    DoCmd.SendObject acSendNoObject, , , strEmail, , , "Store Sales as of " & Now, strSQL, False
ExitSub:
    EOLog "EODeMail Recap" & vbCrLf & CDOmsg.To & vbCrLf & CDOmsg.Subject & vbCrLf & CDOmsg.TextBody
    Exit Sub

ErrLogWithName:

ErrLog:
    EOLog strSQL & vbCrLf & "Error #" & Err.Number & " " & Err.Description & " : EOProcess"
    Resume ExitSub
End Sub


Public Function GetOOG() As String

    GetOOG = vbCrLf & "You can fill this form out and e-mail it back to " & GetKMDSSettings("KMDSMail from", "Agents@kmds.us") & vbCrLf
    GetOOG = GetOOG & vbCrLf & "Don't have excel? You can you use Open Office.." & vbCrLf
    GetOOG = GetOOG & vbCrLf & "Download Open Office from http://www.openoffice.org ." & vbCrLf
    GetOOG = GetOOG & vbCrLf & "The Free and Open Productivity Suite.." & vbCrLf
End Function
