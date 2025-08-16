Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdSetAllowedForThisProgram_Click()
Dim nResp As Integer
Dim stDocName As String
Dim stLinkCriteria As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim rs
    If gcfHandleErrors Then On Error GoTo ErrLog

    nResp = MsgBox("Set ALL Allowed flags to...", vbYesNoCancel)
    Select Case nResp
    Case vbYes
        nResp = -1
    Case vbNo
        nResp = 0
    Case Else
        Exit Sub
    End Select

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCashierAndProgramsSetAllowedFlagByProgram"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parAllowed", adBoolean, adParamInput)
        .Parameters.Append .CreateParameter("parProgramName", adBSTR, adParamInput)
        .Parameters("parProgramName") = cmdSetAllowedForThisProgram.Tag
        .Parameters("parAllowed") = nResp
        .Execute
    End With

    Me.CashierRightsCashierListSub.Requery
    '
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdSetAllowed_Click" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub
