Attribute VB_Name = "modPunchTheClock"
Option Compare Database
Option Explicit

Public Sub PunchTheClock(CashierID As Long, InOrOut As Byte)

'Dim comQDF As New ADODB.Command
'Dim adoCnOLCNetwork As New ADODB.Connection
'
'
'    If gcfHandleErrors Then On Error GoTo ExitLog
'    'added 05/16/2005
'    'to prevent 'catastrophic error'
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        .Properties("Data Source").Value = gstrCashsysInsLoc 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
'        .Properties("Initial Catalog").Value = "KMDSSQL"
'        .Properties("User ID").Value = "sa"
'        .Properties("Password").Value = ""
'        .Open
'    End With
'
'
'    With comQDF
'        .ActiveConnection = adoCnOLCNetwork
'        .CommandText = "sp_TimeClockPunch " & CashierID '& ", " & InOrOut
'        .CommandType = adCmdText
'        .Execute
'    End With
'
'    adoCnOLCNetwork.Close
'
'ExitLog:
'    Set adoCnOLCNetwork = Nothing
'    Set comQDF = Nothing
End Sub

'Public Sub TimeClockPunchAppend()
'
'
''Dim comQDF As New Command
'Dim comQDF As Object
'Set comQDF = CreateObject("ADODB.COMMAND")
'
'
''Dim adoCnOLCNetwork As New ADODB.Connection
'Dim adoCnOLCNetwork As Object
'Set adoCnOLCNetwork = CreateObject("ADODB.Connection")
'
'
'
'    If gcfHandleErrors Then On Error GoTo ErrLog
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        .Properties("Data Source").Value = gstrCashsysInsLoc 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
'        .Properties("Initial Catalog").Value = "KMDSSQL"
'        'integrated security .Properties("User ID").Value = "sa"
'        'integrated security .Properties("Password").Value = ""
'        .ConnectionString = .ConnectionString & "; Trusted_Connection=Yes"
'        .Open
'    End With
'
'
'    With comQDF
'        .ActiveConnection = adoCnOLCNetwork
'        .CommandText = "sp_TCP_Append_Punches"
'        .CommandType = adCmdText
'        .Execute
'    End With
'
'    adoCnOLCNetwork.Close
'
'ExitLog:
'    Set adoCnOLCNetwork = Nothing
'    Set comQDF = Nothing
'    Exit Sub
'
'ErrLog:
'    MsgBox "Error #" & Err.Number & " " & Err.Description & vbcrlf & "TimeClockPunchAppend"
'    Resume ExitLog
'
'End Sub
'
