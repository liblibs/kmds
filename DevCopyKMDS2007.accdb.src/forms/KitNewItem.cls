Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbSalesDepartmentDesc_AfterUpdate()
    cmbSalesDepartmentID = cmbSalesDepartmentDesc

End Sub

Private Sub cmbSalesDepartmentID_AfterUpdate()
    cmbSalesDepartmentDesc = cmbSalesDepartmentID

End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    If txtDescription <> "" Then
        If MsgBox("Save this as a new item?", vbYesNo) = vbYes Then
            SaveThisKit
            DoCmd.Close
        End If

    Else
        MsgBox "Please enter a description before saving!"
        txtDescription.SetFocus
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdfrmSaveRecord_Click"
    Resume ExitSub

End Sub

Private Sub Form_Load()
    Dim rs As Object
    Dim lngID As Long

    SetKMDSSettings "KitLastItem", "NA"
    cmbSalesDepartmentID = GetKMDSSettings("KitSalesDptDflt", "30")
    cmbSalesDepartmentID_AfterUpdate

If gcfHandleErrors Then On Error GoTo ErrLog
    If SalesDepartmentID = 1 Then
        ID.SetFocus
    Else
       'Set RS = CurrentDb().OpenRecordset("select max(ID) as MaxID from Item")
        Set rs = CurrentDb().OpenRecordset("select max(ID) as MaxID from Item where ID < 10000000000")  '07/17/2021
        rs.MoveLast
        ID = rs!MaxID + 1
        'txtUPC.SetFocus
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "SalesDepartmentID_Exit"
    Resume ExitSub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    SetKMDSSettings "KitSalesDptDflt", cmbSalesDepartmentID
End Sub

Private Sub SaveThisKit()


    Dim strSQL As String
    Dim strSize As String
    Dim rs As Object
    Dim nSize As Integer
    Dim nVendorID As Integer


    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNumeric(cmbSalesDepartmentID) Then
        MsgBox ("SalesDepartment ID must have a value!")
        cmbSalesDepartmentID.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(ID) Then
        MsgBox ("ID must have a value!")
        ID.SetFocus
        Exit Sub
    End If

    If IsNull(txtUPC) Or txtUPC = "" Then
        txtUPC = ID
    End If

    If IsNull(txtDescription) Or txtDescription = "" Then
        MsgBox ("Description can not be blank!")
        txtDescription.SetFocus
        Exit Sub
    End If


    'If MsgBox("Save This Item?", vbYesNo, "OK to save?") = vbYes Then
    gbRunMailChanges = True

    strSQL = "INSERT INTO Item ( ID, "
    strSQL = strSQL & "Description, "
    strSQL = strSQL & "SalesDepartmentID) "
    strSQL = strSQL & "SELECT " & ID & " AS ID, "
    strSQL = strSQL & "'" & txtDescription & "' AS Description, "
    strSQL = strSQL & cmbSalesDepartmentID & " AS SalesDepartmentID"

    CurrentDb().Execute (strSQL)

    strSQL = "INSERT INTO UPCItem ( ID, ItemID ) "
    strSQL = strSQL & "SELECT " & ID & " AS ID, "
    strSQL = strSQL & ID & " AS ItemID"
    If gcfHandleErrors Then On Error Resume Next
    CurrentDb().Execute (strSQL)

    strSQL = "INSERT INTO Kit ( ItemID ) "
    strSQL = strSQL & "SELECT " & ID & " AS ItemID "
    If gcfHandleErrors Then On Error Resume Next
    CurrentDb().Execute (strSQL)


    SetKMDSSettings "KitLastItem", ID


'    Else
'        'Cancel = True
'        Exit Sub
'    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "SaveThisKit"
    Resume ExitSub


End Sub
