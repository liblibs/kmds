Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cboCategoryFrom_AfterUpdate()
    cboCategoryTo = cboCategoryFrom
    lstLabels.Requery

End Sub

Private Sub cboCategoryTo_AfterUpdate()
    lstLabels.Requery
End Sub

Private Sub Form_Close()
    If gcfHandleErrors Then On Error GoTo ErrLog
   SetKMDSSettings "LabelListLastSubcategoryID", CStr(cboCategoryFrom)
ExitSub:
    SetOptions True
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Close"
    Resume ExitSub
End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog

    cboCategoryTo = GetKMDSSettings("LabelListLastSubcategoryID", "0")
    cboCategoryFrom = cboCategoryTo
    With lstLabels
        .RowSource = "qryLabelList"

        .Requery
    End With
ExitSub:
    SetOptions True
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    Dim strSaveRowSource As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    lstLabels.RowSource = ""
    'delete work table
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.DeleteObject acTable, "ItemLabels"

    If gcfHandleErrors Then On Error GoTo ErrLog
    'append all items to work table
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelMakeTable"
    End With
    comQDF.Execute
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelMakeTableUpdateLiquorUPC"
    End With
    comQDF.Execute
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

ExitSub:
    SetOptions True
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Open"
    Resume ExitSub
End Sub



Private Sub lstLabels_AfterUpdate()
    Debug.Print lstLabels.ItemData(lstLabels.ListIndex)
    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")
    'set the current record to 'selected'
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelsUpdateLabels"
        .Parameters.Append .CreateParameter("parItemID", adBigInt, adParamInput)
        .Parameters("parItemID") = lstLabels.ItemData(lstLabels.ListIndex)
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    If optShelf Then
        BrotherLabels ("Shelf")
    Else
        BrotherLabels ("Item")
    End If
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "qryItemLabelsResetLabels"
        .Execute
    End With
End Sub
