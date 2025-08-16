Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim bQueryUnload As Boolean

Private Sub chkSkipPhysical_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If chkSkipPhysical Then
        Me.AuditSub("PhysicalCount").TabStop = False
        Me.AuditSub("PhysicalCount").BackColor = cmbSubCategoryID.BackColor

    Else
        Me.AuditSub("PhysicalCount").TabStop = True
        Me.AuditSub("PhysicalCount").BackColor = vbWhite

    End If
    Me.AuditSub.SetFocus
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub

End Sub

Private Sub chkSkipVariance_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If chkSkipVariance Then
        Me.AuditSub("VarianceAmount").TabStop = False
        Me.AuditSub("VarianceAmount").BackColor = cmbSubCategoryID.BackColor

    Else
        Me.AuditSub("VarianceAmount").TabStop = True
        Me.AuditSub("VarianceAmount").BackColor = vbWhite

    End If
    Me.AuditSub.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub

End Sub

Private Sub cmbListType_AfterUpdate()

Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
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

    Case "Size"
        strSQL = "SELECT Description, ID, ID "
        strSQL = strSQL & "FROM Size "
        strSQL = strSQL & "ORDER BY ID"
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
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub

End Sub

Private Sub cmbSubCategoryID_AfterUpdate()
    If gcfHandleErrors Then On Error Resume Next
    cmdAddItems.Visible = True

    cmbSubCategoryIDEnd = cmbSubCategoryID
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub
End Sub




Private Sub cmbSubCategoryIDEnd_AfterUpdate()
    If gcfHandleErrors Then On Error Resume Next
    cmdAddItems.Visible = True

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub

End Sub

Private Sub cmdAddItems_Click()
Dim strTemp As String
Dim varReturn As Variant
    If gcfHandleErrors Then On Error GoTo ErrLog
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
    Case "Size"
        strTemp = "qryAuditAppendBySize"

    Case Else
    End Select
    SetOptions False
    'CurrentDb.Execute strTemp, False

    DoCmd.OpenQuery strTemp, acNormal, acEdit
    SetOptions True
    Me.Refresh
    Me.Requery

    Me.AuditSub.SetFocus
    cmdAddItems.Visible = False
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub
End Sub

Private Sub cmdAuditUpdateOnShelf_Click()
    Dim rs As Object
    Dim nCode As Double '07/17/2021
    If gcfHandleErrors Then On Error GoTo ErrLog
    With Me.AuditSub.Form
        nCode = .ItemID
        AuditUpdateOnShelf
        'If gcfHandleErrors Then On Error GoTo ErrLog
        Set rs = .RecordsetClone ' Me.Recordset.Clone
        rs.FindFirst "[ItemID] = " & Str(nCode)
        .Bookmark = rs.Bookmark

    End With
    Me.AuditSub.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub
End Sub

Private Sub cmdfrmClose_Click()
    DoCmd.Close
End Sub


Private Sub Command21_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("Delete all records?" & vbCrLf & "Are you sure?", vbYesNo) = vbYes Then
        SetOptions False
        DoCmd.RunSQL "Delete * from Audit", False
        SetOptions True
        Me.Refresh
    End If
    Me.AuditSub.SetFocus

    bQueryUnload = False
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub
End Sub

Private Sub cmdPost_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPost_Click

    Dim stDocName As String

    If MsgBox("Apply variances to current onhands?", vbYesNo, "POST THIS BATCH?") = vbYes Then
        chkOnlyVariance = False
        ArchiveReport "Audit", glsCashierName
        chkOnlyVariance = True
        cmdPrint_Click

        bQueryUnload = False
        DoCmd.Close



        If AuditBatchPost Then
            MsgBox "Physical Counts have been applied!"
        End If
    End If
Exit_cmdPost_Click:
    Exit Sub

Err_cmdPost_Click:
    MsgBox Err.Description
    Resume Exit_cmdPost_Click

End Sub


Private Sub Command52_Click()

End Sub

Private Sub Form_Load()
    AuditUpdateOnShelf
    SetCountByOptions

    Me.InsideHeight = Box80.Height + FormFooter.Height + FormHeader.Height + 100
    Me.InsideWidth = Box80.Width + 500
    If gbPalmCommaDelimited Then
        Label57.Caption = "Receive from Palm (Comma)"
    Else
        Label57.Caption = "Receive from Palm (Tab)"
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        SetFormSize Me

    End If
    bQueryUnload = True
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMsg As String
    If gcfHandleErrors Then On Error GoTo ErrLog
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
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub
End Sub

Private Sub Frame24_AfterUpdate()

End Sub
Private Sub cmdPrint_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrint_Click
    Dim strOrderBy As String
    Dim stDocName As String
    strOrderBy = ""
     Select Case framSortBy.Value
    Case 1
        strOrderBy = "ItemID"
    Case 2
        strOrderBy = "Description,SizeDesc"
    Case 3
        strOrderBy = "SizeDesc,Description"
    Case 4
        strOrderBy = "SetupDate"
    End Select


    If Not gcfHandleErrors Then Debug.Print Me.AuditSub.Form.OrderBy
    stDocName = "Audit"
    DoCmd.OpenReport stDocName, acViewNormal, , , , strOrderBy

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

Private Sub cmdPreview_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPreview_Click
    Dim strOrderBy As String
    Dim stDocName As String
    strOrderBy = ""
     Select Case framSortBy.Value
    Case 1
        strOrderBy = "ItemID"
    Case 2
        strOrderBy = "Description,SizeDesc"
    Case 3
        strOrderBy = "SizeDesc,Description"
    Case 4
        strOrderBy = "SetupDate"
    End Select


    If Not gcfHandleErrors Then Debug.Print Me.AuditSub.Form.OrderBy
    stDocName = "Audit"
    DoCmd.OpenReport stDocName, acPreview, , , , strOrderBy

Exit_cmdPreview_Click:
    Exit Sub

Err_cmdPreview_Click:
    Select Case Err.Number
    Case 2501
        'canceled because of nothing to print!
    Case Else
        MsgBox Err.Number & " " & Err.Description
    End Select
    Resume Exit_cmdPreview_Click

End Sub

Private Sub AuditUpdateOnShelf()
If gcfHandleErrors Then On Error GoTo Err_cmdAuditUpdateOnShelf_Click

    Dim stDocName As String
    'changed 2008-12-13 to use stored procedure
    stDocName = "sp_AuditUpdateOnShelf"
    SQLPassThru stDocName
    Me.AuditSub.Requery
Exit_cmdAuditUpdateOnShelf_Click:
    Exit Sub

Err_cmdAuditUpdateOnShelf_Click:
    MsgBox Err.Description
    Resume Exit_cmdAuditUpdateOnShelf_Click

End Sub

Private Sub framSortBy_AfterUpdate()
    Dim nCode As Double '07/17/2021
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    nCode = Me.AuditSub.Form.ItemID

    Select Case framSortBy.Value
    Case 1
        Me.AuditSub.Form.OrderBy = "ItemID"
    Case 2
        Me.AuditSub.Form.OrderBy = "Lookup_ItemDescription.Description,Lookup_Size.PkgSize"
    Case 3
        Me.AuditSub.Form.OrderBy = "Lookup_Size.PkgSize,Lookup_ItemDescription.Description"
    Case 4
        Me.AuditSub.Form.OrderBy = "SetupDate"
    End Select

    Me.AuditSub.Form.Refresh
    With Me.AuditSub.Form
         'If gcfHandleErrors Then On Error GoTo ErrLog
        Set rs = .RecordsetClone ' Me.Recordset.Clone
        rs.FindFirst "[ItemID] = " & Str(nCode)
        .Bookmark = rs.Bookmark

    End With
    Me.AuditSub.SetFocus
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Error!"
    Resume ExitSub

End Sub
Private Sub cmdPalm_Click()
    Dim fname As Variant
    fname = gstrPalmDataDirectoryInventoryTo
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.DeleteObject acTable, "InvoLineUPCMaxDate"
    DoCmd.DeleteObject acTable, "PalmTo"

    If gcfHandleErrors Then On Error GoTo Err_cmdPalm_Click

    Dim stDocName As String
    SetOptions False
    stDocName = "qryAuditFileToPalm_Step_1"
   ' DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryAuditFileToPalm_Step_2"
   ' DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryAuditFileToPalm_Step_3"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    DoCmd.TransferText acExportDelim, "Palm Export Specification", "PalmTo", gstrPalmDataDirectoryInventoryTo, False

    SetOptions True

Exit_cmdPalm_Click:
    Exit Sub

Err_cmdPalm_Click:
    MsgBox Err.Description
    Resume Exit_cmdPalm_Click

End Sub
Private Sub cmdPalmFrom_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPalmFrom_Click

    Dim stDocName As String

    SetOptions False
    stDocName = "qryAuditFileFromPalm_Step_1"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    If gbPalmCommaDelimited Then
        stDocName = "PalmFrom Import Specification Comma"
    Else
        stDocName = "PalmFrom Import Specification Tab" 'there is an extra spacen in this one..
    End If

    DoCmd.TransferText acImportDelim, stDocName, "PalmFrom", gstrPalmDataDirectoryInventoryFrom, False

    stDocName = "qryAuditFileFromPalm_Step_2"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryAuditFileFromPalm_Step_3"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryAuditFileFromPalm_Step_4"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    'Shell ("del " & gstrPalmDataDirectoryInventoryFrom)
    Kill gstrPalmDataDirectoryInventoryFrom

    Me.Requery
    cmdAuditUpdateOnShelf_Click
    SetOptions True

Exit_cmdPalmFrom_Click:
    Exit Sub

Err_cmdPalmFrom_Click:
    MsgBox Err.Description
    Resume Exit_cmdPalmFrom_Click

End Sub
Private Sub cmdReassign_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdReassign_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.Close

    stDocName = "ReAssignItemID"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdReassign_Click:
    Exit Sub

Err_cmdReassign_Click:
    MsgBox Err.Description
    Resume Exit_cmdReassign_Click

End Sub

Private Sub optCountBy_AfterUpdate()
    SetCountByOptions
End Sub

Private Sub SetCountByOptions()

    Me.AuditSub("PhysicalCount").ColumnHidden = False
    Me.AuditSub("VarianceAmount").ColumnHidden = False

    Me.AuditSub("Front").ColumnHidden = True
    Me.AuditSub("Back").ColumnHidden = True
    Me.AuditSub("Display").ColumnHidden = True
    Me.AuditSub("Overflow").ColumnHidden = True
    optLocations.Visible = False

    Me.AuditSub.SetFocus
    Select Case optCountBy.Value
    Case 1 '
            chkSkipPhysical = False
            chkSkipVariance = True
            Me.AuditSub("PhysicalCount").SetFocus


    Case 2 '
            chkSkipPhysical = True
            chkSkipVariance = False

            Me.AuditSub("VarianceAmount").SetFocus

    Case Else '
            chkSkipPhysical = True
            chkSkipVariance = True
            Me.AuditSub("Front").ColumnHidden = False
            Me.AuditSub("Back").ColumnHidden = False
            Me.AuditSub("Display").ColumnHidden = False
            Me.AuditSub("Overflow").ColumnHidden = False
            'Me.AuditSub("PhysicalCount").ColumnHidden = True
            optLocations.Visible = True
            SetLocationOptions
    End Select
    chkSkipPhysical_AfterUpdate
    chkSkipVariance_AfterUpdate
End Sub

Public Sub SetLocationOptions()
    Me.AuditSub("Front").TabStop = False
    Me.AuditSub("Back").TabStop = False
    Me.AuditSub("Display").TabStop = False
    Me.AuditSub("Overflow").TabStop = False
    Me.AuditSub.SetFocus

    Select Case optLocations.Value
    Case 1 ' front
        Me.AuditSub("Front").TabStop = True
        Me.AuditSub("Front").SetFocus
    Case 2 'back
        Me.AuditSub("Back").TabStop = True
        Me.AuditSub("Back").SetFocus
    Case 3 'display
        Me.AuditSub("Display").TabStop = True
        Me.AuditSub("Display").SetFocus
    Case Else ' Overflow
        Me.AuditSub("Overflow").TabStop = True
        Me.AuditSub("Overflow").SetFocus
    End Select

End Sub

Private Sub optLocations_AfterUpdate()
    SetLocationOptions
End Sub

Private Sub cmdAddNewRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNewRecord_Click
    Me.AuditSub.SetFocus


    DoCmd.GoToRecord , , acNewRec


Exit_cmdAddNewRecord_Click:
    Exit Sub

Err_cmdAddNewRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNewRecord_Click

End Sub
