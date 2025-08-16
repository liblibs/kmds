Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'    Select Case UCase(Combo8)
'    Case "SALESDEPARTMENT"
'    Case "CATEGORY"
'    Case "PRICE GROUP"
'    Case "VENDOR"
'    Case "SALES LEADS"
'    Case Else
'        MsgBox Combo8 & " not programmed"
'        Exit Sub
'    End Select

Private Sub chkFilterLiquor_AfterUpdate()
    Combo8_AfterUpdate
End Sub

Private Sub chkFilterRelated_AfterUpdate()
    Combo8_AfterUpdate
End Sub

Private Sub chkFilterRelated_Click()
    Combo8_AfterUpdate
End Sub

Private Sub cmdClearSelections_Click()
Dim nI As Integer
Dim nJ As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    'whittle it down!
    nJ = lstChildNo.ItemsSelected.Count - 1
    For nI = 0 To nJ
        lstChildNo.Selected(lstChildNo.ItemsSelected(0)) = False
    Next nI
    lstChildNo.SetFocus
    cmdMultiSelect.Visible = False
    cmdClearSelections.Visible = False

ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cmdClearSelections_Click"
    Resume ExitSub
End Sub

Private Sub cmdMultiSelect_Click()

Dim nI As Integer
Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    Select Case UCase(Combo8)
    Case "SALESDEPARTMENT"
        If lstParent > 1 Then

            For nI = 0 To lstChildNo.ItemsSelected.Count - 1
                strSQL = ""
                strSQL = "UPDATE Item "
                strSQL = strSQL & "SET Item.SalesDepartmentID = " & CLng(lstParent)
                strSQL = strSQL & " WHERE (((Item.ID)=" & lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI)) & "))"
                DoCmd.RunSQL strSQL
            Next nI
        End If

    Case "CATEGORY"
        For nI = 0 To lstChildNo.ItemsSelected.Count - 1
            strSQL = ""
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.SubCategoryID = " & CLng(lstParent)
            'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI))) & "))"
            strSQL = strSQL & " WHERE (((Item.ID)=" & lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI)) & "))"
            DoCmd.RunSQL strSQL
        Next nI
    Case "PRICE GROUP"
            For nI = 0 To lstChildNo.ItemsSelected.Count - 1
            strSQL = ""
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.PriceGroupID = " & CLng(lstParent)
            strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI))) & "))"
            DoCmd.RunSQL strSQL
        Next nI
    Case "VENDOR"
        For nI = 0 To lstChildNo.ItemsSelected.Count - 1
            strSQL = ""
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.VendorID = " & CLng(lstParent)
            strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI))) & "))"
            DoCmd.RunSQL strSQL
        Next nI
    Case "SALES LEADS"
        For nI = 0 To lstChildNo.ItemsSelected.Count - 1
            strSQL = ""
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.SalesLeadID = " & CLng(lstParent)
            strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI))) & "))"
            DoCmd.RunSQL strSQL
        Next nI
    Case "PICTURES"
        For nI = 0 To lstChildNo.ItemsSelected.Count - 1
            strSQL = ""
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.PictureID = " & CLng(lstParent)
            strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ItemsSelected.Item(nI))) & "))"
            DoCmd.RunSQL strSQL
        Next nI
    Case Else
        MsgBox Combo8 & " not programmed"
        Exit Sub
    End Select
    SetOptions True
    cmdClearSelections_Click
    RequeryLists
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdMultiSelect_Click"
    Resume ExitSub

End Sub

Private Sub Combo8_AfterUpdate()
    Dim strSQLParent As String
    Dim strSQLChild As String
    Dim strSQLChildNo As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    imgPicture.Visible = False
    lstParent.Height = lstChild.Height


    Select Case UCase(Combo8)

    Case "SALESDEPARTMENT"
        chkFilterLiquor = True
        chkFilterRelated = False
        strSQLParent = "SELECT [SalesDepartment].[ID], [SalesDepartment].[Description] "
        strSQLParent = strSQLParent & "FROM [SalesDepartment] WHERE ID <> 1 ORDER BY [ID]"

        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.SalesDepartmentID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.SalesDepartmentID)<>clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID<> 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.SalesDepartmentID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.SalesDepartmentID)=clng([Forms]![ItemsByListBox]![lstParent])) "
        If chkFilterLiquor Then
            strSQLChild = strSQLChild & " AND (Item.SalesDepartmentID<> 1)"
        End If

        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"

    Case "CATEGORY"
        chkFilterLiquor = False
'        chkFilterRelated = False
        strSQLParent = "SELECT SubCategory.ID, SubCategory.Description FROM SubCategory "
        strSQLParent = strSQLParent & "ORDER BY  SubCategory.CategoryID, SubCategory.Description;"

        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.SubCategoryID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.SubCategoryID)<>clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterRelated Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID = 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.SubCategoryID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.SubCategoryID)=clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterRelated Then
            strSQLChild = strSQLChild & " AND (Item.SalesDepartmentID = 1)"
        End If
        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"


    Case "PRICE GROUP"
        chkFilterLiquor = True
        chkFilterRelated = False
        strSQLParent = "SELECT PriceGroup.ID, PriceGroup.Description FROM PriceGroup ORDER BY PriceGroup.Description;"

        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.PriceGroupID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.PriceGroupID)<>clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID<> 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.PriceGroupID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.PriceGroupID)=clng([Forms]![ItemsByListBox]![lstParent])) "

        If chkFilterLiquor Then
            strSQLChild = strSQLChild & " AND (Item.SalesDepartmentID<> 1)"
        End If

        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"



    Case "VENDOR"
        chkFilterLiquor = True
        chkFilterRelated = False

        strSQLParent = "SELECT Vendor.ID, Vendor.Name FROM Vendor WHERE ID <> 1 ORDER BY Vendor.Name;"

        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.VendorID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.VendorID)<> clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID<> 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.VendorID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.VendorID)=clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChild = strSQLChild & " AND (Item.SalesDepartmentID<> 1)"
        End If
        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"
     Case "SALES LEADS"
        'chkFilterLiquor = True
        'chkFilterRelated = False

        strSQLParent = "SELECT SalesLeads.ID, SalesLeads.Text From SalesLeads " 'WHERE (((SalesLeads.SpecificItemOnly) = True))"


        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.SalesLeadID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.SalesLeadID)<> clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID <> 1)"
        End If
        If chkFilterRelated Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID = 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.SalesLeadID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.SalesLeadID)=clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChild = strSQLChild & " AND (Item.SalesDepartmentID<> 1)"
        End If
        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"
     Case "PICTURES"
        'chkFilterLiquor = True
        'chkFilterRelated = False
        lstParent.Height = imgPicture.Height
        imgPicture.Visible = True

        strSQLParent = "SELECT SalesPictures.ID, SalesPictures.FileName From SalesPictures" ' WHERE (((Pictures.SpecificItemOnly) = True))"


        strSQLChildNo = "SELECT Item.ID, Item.Description, Item.Price, Item.PictureID "
        strSQLChildNo = strSQLChildNo & "FROM Item WHERE (((Item.PictureID)<> clng([Forms]![ItemsByListBox]![lstParent]))"
        If chkFilterLiquor Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID <> 1)"
        End If
        If chkFilterRelated Then
            strSQLChildNo = strSQLChildNo & " AND (Item.SalesDepartmentID = 1)"
        End If

        strSQLChildNo = strSQLChildNo & ") ORDER BY Item.Description;"

        strSQLChild = "SELECT Item.ID, Item.Description, Item.Price, Item.PictureID "
        strSQLChild = strSQLChild & "FROM Item WHERE (((Item.PictureID)=clng([Forms]![ItemsByListBox]![lstParent]))"

        strSQLChild = strSQLChild & ") ORDER BY Item.Description;"
    Case Else
        MsgBox Combo8 & " not programmed"
        Exit Sub
    End Select
    lblParent.Caption = Combo8

    lstParent.RowSource = strSQLParent
    lstChildNo.RowSource = strSQLChildNo
    lstChild.RowSource = strSQLChild
    cmbParentPrevious.RowSource = strSQLParent

    lstParent.Requery
    cmbParentPrevious.Requery

    RequeryLists
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Combo8_AfterUpdate"
    Resume ExitSub
End Sub





Private Sub Form_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    imgPicture.Visible = False
    lstParent.Height = lstChild.Height
    lstParent.RowSource = ""
    lstChildNo.RowSource = ""
    lstChild.RowSource = ""
    cmbParentPrevious.RowSource = ""

    lstParent.Requery
    cmbParentPrevious.Requery
    Combo8 = "SalesDepartment"
    Combo8_AfterUpdate

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub lstChild_DblClick(Cancel As Integer)
    Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    If IsNull(cmbParentPrevious) Then
        MsgBox "You must first select a target from the list below!"
        cmbParentPrevious.SetFocus
        Exit Sub
    End If
    If (cmbParentPrevious) = lstParent Then
        MsgBox "You must first select a DIFFERENT target from the list below!"
        cmbParentPrevious.SetFocus
        Exit Sub
    End If
    Select Case UCase(Combo8)
    Case "SALESDEPARTMENT"
        If cmbParentPrevious > 1 Then
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.SalesDepartmentID = " & CLng(cmbParentPrevious)
            'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
            strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
            DoCmd.RunSQL strSQL, False
        End If
    Case "CATEGORY"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.SubCategoryID = " & CLng(cmbParentPrevious)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False
    Case "PRICE GROUP"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.PriceGroupID = " & CLng(cmbParentPrevious)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False
    Case "VENDOR"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.VendorID = " & CLng(cmbParentPrevious)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False
    Case "SALES LEADS"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.SalesLeadID = " & CLng(cmbParentPrevious)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False
    Case "PICTURES"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.PictureID = " & CLng(cmbParentPrevious)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChild.ItemData(lstChild.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChild.ItemData(lstChild.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False

    Case Else
        MsgBox Combo8 & " not programmed"
        Exit Sub
    End Select

    RequeryLists
    SetOptions True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdMultiSelect_Click"
    Resume ExitSub

End Sub

Private Sub lstChildNo_Click()

    cmdMultiSelect.Visible = True
    cmdClearSelections.Visible = True
    cmbParentPrevious = lstChildNo.Column(3)
End Sub

Private Sub lstChildNo_DblClick(Cancel As Integer)
    Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    Select Case UCase(Combo8)
    Case "SALESDEPARTMENT"
        If lstParent > 1 Then
            strSQL = "UPDATE Item "
            strSQL = strSQL & "SET Item.SalesDepartmentID = " & CLng(lstParent)
            strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
            DoCmd.RunSQL strSQL, False
        Else
            Beep
        End If
    Case "CATEGORY"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.SubCategoryID = " & CLng(lstParent)
        'strsql = strsql & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
        strSQL = strSQL & " WHERE (((Item.ID)=" & lstChildNo.ItemData(lstChildNo.ListIndex) & "))"
        DoCmd.RunSQL strSQL, False
    Case "PRICE GROUP"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.PriceGroupID = " & CLng(lstParent)
        strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
        DoCmd.RunSQL strSQL, False
    Case "VENDOR"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.VendorId = " & CLng(lstParent)
        strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
        DoCmd.RunSQL strSQL, False
    Case "SALES LEADS"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.SalesLeadId = " & CLng(lstParent)
        strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
        DoCmd.RunSQL strSQL, False
    Case "PICTURES"
        strSQL = "UPDATE Item "
        strSQL = strSQL & "SET Item.PictureID = " & CLng(lstParent)
        strSQL = strSQL & " WHERE (((Item.ID)=" & CLng(lstChildNo.ItemData(lstChildNo.ListIndex)) & "))"
        DoCmd.RunSQL strSQL, False
    Case Else
        MsgBox Combo8 & " not programmed"
        Exit Sub
    End Select
    RequeryLists
    SetOptions True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdMultiSelect_Click"
    Resume ExitSub

End Sub

Private Sub lstParent_AfterUpdate()
    RequeryLists

End Sub

Private Sub RequeryLists()
    lstChild.Requery
    lstChildNo.Requery
    lblChild.Caption = lstChild.ListCount & " Items In " & Combo8 & " = " & lstParent.Column(1)
    lblChildNo.Caption = lstChildNo.ListCount & " Items NOT In " & Combo8 & " = " & lstParent.Column(1)
    lblParent.Caption = Combo8 & " (" & (lstChild.ListCount + lstChildNo.ListCount) & ") items"

    Select Case UCase(Combo8)
    Case "PICTURES"
        If Not IsNull(lstParent.Column(1, lstParent.ListIndex)) Then
        If Dir(Trim(lstParent.Column(1, lstParent.ListIndex))) <> "" Then
            imgPicture.Picture = lstParent.Column(1, lstParent.ListIndex)
        End If
        End If
    Case Else
    End Select
End Sub
Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
