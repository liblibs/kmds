Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bQueryUnload As Boolean

Private Sub cmbListType_AfterUpdate()

Dim strSQL As String
    Description_Label.Caption = cmbListType
    cmdAddItems.Caption = "Add items from this range of " & cmbListType & "s."
    cmdAddItems.Visible = True

    Select Case cmbListType
    Case "Code"
        strSQL = "SELECT Item.ID, Item.Description, Item.ID "
        strSQL = strSQL & "FROM Item "
        strSQL = strSQL & "ORDER BY len(Item.ID),Item.ID;"
        cmbSubCategoryID.ColumnWidths = "720,1440,0"

    Case "Description"
        strSQL = "SELECT Item.Description, Item.ID, Item.ID "
        strSQL = strSQL & "FROM Item "
        strSQL = strSQL & "ORDER BY Item.Description;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case "SalesDept"
        strSQL = "SELECT SalesDepartment.Description, SalesDepartment.ID, SalesDepartment.ID "
        strSQL = strSQL & "FROM SalesDepartment "
        strSQL = strSQL & "ORDER BY SalesDepartment.Description;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case "Vendor"
        strSQL = "SELECT Vendor.Name, Vendor.ID, Vendor.ID "
        strSQL = strSQL & "FROM Vendor "
        strSQL = strSQL & "ORDER BY Vendor.Name;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case "PriceGroup"
        strSQL = "SELECT PriceGroup.Description, PriceGroup.ID, PriceGroup.ID "
        strSQL = strSQL & "FROM PriceGroup "
        strSQL = strSQL & "ORDER BY PriceGroup.Description;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case "Location"
        strSQL = "SELECT Item.Location, Item.ID, Item.ID "
        strSQL = strSQL & "FROM Item "
        strSQL = strSQL & "ORDER BY Item.Location;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case "Category"
        strSQL = "SELECT SubCategory.Description, SubCategory.ID, SubCategory.ID "
        strSQL = strSQL & "FROM SubCategory "
        strSQL = strSQL & "ORDER BY SubCategory.Description;"
        cmbSubCategoryID.ColumnWidths = "1440,720,0"

    Case Else
    End Select

    cmbSubCategoryID.RowSource = strSQL
    cmbSubCategoryID.Requery

    cmbSubCategoryID = ""

    cmbSubCategoryIDEnd.BoundColumn = cmbSubCategoryID.BoundColumn
    cmbSubCategoryIDEnd.ColumnWidths = cmbSubCategoryID.ColumnWidths
    cmbSubCategoryIDEnd.RowSource = cmbSubCategoryID.RowSource
    cmbSubCategoryIDEnd.Requery
    cmbSubCategoryIDEnd = ""

End Sub

Private Sub cmbSubCategoryID_AfterUpdate()
    If gcfHandleErrors Then On Error Resume Next
    cmbSubCategoryIDEnd = cmbSubCategoryID
End Sub




Private Sub cmdAddItems_Click()
Dim strTemp As String
Dim varReturn As Variant
    Select Case cmbListType
    Case "Code"
        strTemp = "qryAuditAppendByCode"
    Case "Description"
        strTemp = "qryAuditAppendByDescription"
    Case "SalesDept"
        strTemp = "qryAuditAppendBySalesDept"
    Case "Vendor"
        strTemp = "qryAuditAppendByVendor"
    Case "PriceGroup"
        strTemp = "qryAuditAppendByPriceGroup"
    Case "Location"
        strTemp = "qryAuditAppendByLocation"
    Case "Category"
        strTemp = "qryAuditAppendByCategory"

    Case Else
    End Select
    SetOptions False
    'CurrentDb.Execute strTemp, False

    DoCmd.OpenQuery strTemp, acNormal, acEdit
    SetOptions True
    Me.Refresh
    Me.Requery
End Sub

Private Sub cmdAuditUpdateOnShelf_Click()
    AuditUpdateOnShelf
End Sub

Private Sub cmdfrmClose_Click()
    DoCmd.Close
End Sub

Private Sub Command21_Click()
    If MsgBox("Delete all records?" & vbCrLf & "Are you sure?", vbYesNo) = vbYes Then
        SetOptions False
        DoCmd.RunSQL "Delete * from Audit", False
        SetOptions True
        Me.Refresh
    End If
    bQueryUnload = False
End Sub

Private Sub cmdPost_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPost_Click

    Dim stDocName As String
    ArchiveReport "Audit", glsCashierName
    bQueryUnload = False
    DoCmd.Close
    If AuditBatchPost Then
        MsgBox "Physical Counts have been applied!"
    End If
Exit_cmdPost_Click:
    Exit Sub

Err_cmdPost_Click:
    MsgBox Err.Description
    Resume Exit_cmdPost_Click

End Sub


Private Sub Form_Load()
    AuditUpdateOnShelf
End Sub

Private Sub Form_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
    End If
    bQueryUnload = True
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMsg As String
    strMsg = "This batch has not been posted!" & vbCrLf
    strMsg = strMsg & "To apply the physical counts you must click the 'Post This Batch' Button" & vbCrLf & vbCrLf
    If bQueryUnload = True Then

        'If MsgBox(strMsg, vbYesNo) = vbYes Then
        '    bQueryUnload = False
        'Else
        '    bQueryUnload = True
        '    Cancel = True
        'End If
    End If
End Sub

Private Sub Frame24_AfterUpdate()
    Select Case Frame24.Value
    Case 1
        Me.AuditSub.Form.OrderBy = "Lookup_cmbItemID.ID"
    Case 2
        Me.AuditSub.Form.OrderBy = "Lookup_ItemDescription.Description"
    Case 3
        Me.AuditSub.Form.OrderBy = "Lookup_Size.PkgSize,Lookup_ItemDescription.Description"
    End Select
    Me.AuditSub.Form.Refresh
End Sub
Private Sub cmdPrint_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrint_Click
    Dim strOrderBy As String
    Dim stDocName As String
    strOrderBy = ""
     Select Case Frame24.Value
    Case 1
        strOrderBy = "ItemID"
    Case 2
        strOrderBy = "Description"
    Case 3
        strOrderBy = "SizeDesc,Description"
    End Select


    If Not gcfHandleErrors Then Debug.Print Me.AuditSub.Form.OrderBy
    stDocName = "Audit"
    DoCmd.OpenReport stDocName, acPreview, , , , strOrderBy

Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    Select Case Err.Number
    Case 2501
        'canceled because of nothing to print!
    Case Else
        MsgBox Err.Number & " " & Err.Description
    End Select
    Resume Exit_cmdPrint_Click

End Sub

Private Sub AuditUpdateOnShelf()
If gcfHandleErrors Then On Error GoTo Err_cmdAuditUpdateOnShelf_Click

    Dim stDocName As String

    stDocName = "qryAuditUpdateOnShelf"
    SetOptions False
    DoCmd.OpenQuery stDocName, acNormal, acEdit

    SetOptions False
    Me.AuditSub.Requery
Exit_cmdAuditUpdateOnShelf_Click:
    Exit Sub

Err_cmdAuditUpdateOnShelf_Click:
    MsgBox Err.Description
    Resume Exit_cmdAuditUpdateOnShelf_Click

End Sub
