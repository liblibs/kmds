Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbItemByDescription_AfterUpdate()
    Combo1 = cmbItemByDescription

End Sub

Private Sub Combo1_AfterUpdate()
    cmbItemByDescription = Combo1
End Sub
Private Sub cmdChange_Click()


'Dim adoCnOLCNetwork As New ADODB.Connection
Dim adoCnOLCNetwork As Object
Set adoCnOLCNetwork = CreateObject("ADODB.Connection")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog
'    'added 05/16/2005
'    'to prevent 'catastrophic error'
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        .Properties("Data Source").Value = gstrCashsysInsLoc ' "KMDSSQL" 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
'        .Properties("Initial Catalog").Value = "KMDSSQL7"
'        'integrated security .Properties("User ID").Value = "sa"
'        'integrated security .Properties("Password").Value = ""
'        .ConnectionString = .ConnectionString & "; Trusted_Connection=Yes"
'        .Open
'    End With
   '4/11/2014
    With adoCnOLCNetwork
        .ConnectionTimeout = 10
        .ConnectionString = "Provider=SQLNCLI.1;" _
        & "SERVER=" & gstrCashsysInsLoc & ";" _
        & "Database=KMDSSQL7;" _
        & "DataTypeCompatibility=80;" _
        & "Trusted_Connection=Yes;"
        .Open
    End With



    With comQDF
        .ActiveConnection = adoCnOLCNetwork
        .CommandText = "sp_ItemIDReassign " & Combo1 & ", " & CLng(txtNewId)
        .CommandType = adCmdText
        '.Parameters.Append .CreateParameter("parOldItemID", adInteger, adParamInput)
        '.Parameters("parOldItemID") = Combo1
        '.Parameters.Append .CreateParameter("parNewItemID", adInteger, adParamInput)
        '.Parameters("parNewItemID") = CLng(txtNewId)

        'gives 'duplicate records not allowed after first punch
        'If gcfHandleErrors Then On Error Resume Next
        .Execute
        If gcfHandleErrors Then On Error GoTo ExitLog
    End With
    Set comQDF = Nothing

    adoCnOLCNetwork.Close
    Combo1.Requery
    cmbItemByDescription.Requery
    Combo1 = txtNewId
    cmbItemByDescription = Combo1



ExitLog:
    Set adoCnOLCNetwork = Nothing
    Set comQDF = Nothing
    Exit Sub

ErrLog:
    MsgBox Err.Number & " " & Err.Description
    Resume ExitLog
End Sub
Private Sub cmdUPCFix_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdUPCFix_Click

    Dim stDocName As String
    SetOptions (False)
    stDocName = "qryUPCItemsDeleteLen4to7"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    AppendUPCItemsWithItemID
    SetOptions (True)

Exit_cmdUPCFix_Click:
    Exit Sub

Err_cmdUPCFix_Click:
    MsgBox Err.Description
    Resume Exit_cmdUPCFix_Click

End Sub
