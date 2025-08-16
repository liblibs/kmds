Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Combo11_AfterUpdate()
End Sub

Private Sub cmbBusinessDate_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[BusinessDate] = #" & Format(Me![cmbBusinessDate], "mm\/dd\/yyyy") & "#"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark

End Sub

Private Sub Form_Current()
    cmbBusinessDate = BusinessDate
End Sub
Private Sub Command20_Click()
If gcfHandleErrors Then On Error GoTo Err_Command20_Click

    Dim stDocName As String
    stDocName = "LotteryScratchTicketRecon"
    ArchiveReport stDocName
    DoCmd.OpenReport stDocName, acViewNormal

Exit_Command20_Click:
    Exit Sub

Err_Command20_Click:
    MsgBox Err.Description
    Resume Exit_Command20_Click

End Sub

Private Sub Form_Load()
    Form.InsideHeight = Box22.Height + 900
    Form.InsideWidth = Box22.Width + 850

End Sub
Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub Form_Unload(Cancel As Integer)


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


'Dim adoCnOLCNetwork As New ADODB.Connection
Dim adoCnOLCNetwork As Object
Set adoCnOLCNetwork = CreateObject("ADODB.Connection")

    'goto to first record and record the readings in the master tables 'last read' field

    If gcfHandleErrors Then On Error GoTo ErrLog
    DoCmd.GoToRecord , , acFirst
'    'added 05/16/2005
'    'to prevent 'catastrophic error'
'    With adoCnOLCNetwork
'        .ConnectionTimeout = 10
'        .Provider = "sqloledb"
'        .Properties("Data Source").Value = gstrCashsysInsLoc 'adoCnOLCNetworkg adoCnOLCNetworkrsSysfile!CashsysInsLoc ' "KMDSSQL"
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
        .CommandText = "sp_LotteryScratch_UpdateLastRead '" & CStr(Format(cmbBusinessDate, "mm/dd/yyyy")) & "'"
        .CommandType = adCmdText
        .Execute
    End With

    adoCnOLCNetwork.Close

ExitLog:
    Set adoCnOLCNetwork = Nothing
    Set comQDF = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error" & Err.Number & " " & Err.Description
    Resume ExitLog

End Sub
Private Sub cmdSetupNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSetupNew_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "LotteryScratchSetup"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_cmdSetupNew_Click:
    Exit Sub

Err_cmdSetupNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetupNew_Click

End Sub
Private Sub Command23_Click()
If gcfHandleErrors Then On Error GoTo Err_Command23_Click


    DoCmd.Close

Exit_Command23_Click:
    Exit Sub

Err_Command23_Click:
    MsgBox Err.Description
    Resume Exit_Command23_Click

End Sub
Private Sub Command26_Click()
If gcfHandleErrors Then On Error GoTo Err_Command26_Click


    DoCmd.GoToRecord , , acFirst

Exit_Command26_Click:
    Exit Sub

Err_Command26_Click:
    MsgBox Err.Description
    Resume Exit_Command26_Click

End Sub
