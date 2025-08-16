Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim gstDocName As String
Dim gstLinkCriteria As String
Dim gstrInvoiceNo As String
Dim bFintech As Boolean
Dim bFintechLiquorOnly As Boolean


Private Sub chkCategory_Click()
    ShowColumns
End Sub

Private Sub chkCost_Click()
    ShowColumns
End Sub

Private Sub chkDesc_Click()
    ShowColumns
End Sub

Private Sub chkExtCost_Click()
    ShowColumns
End Sub

Private Sub chkExtPrice_Click()
    ShowColumns
End Sub

Private Sub chkItemID_Click()
    ShowColumns
End Sub

Private Sub chkLineID_Click()
    ShowColumns
End Sub

Private Sub chkPrice_click()
    ShowColumns
End Sub

Private Sub chkQuantity_Click()
    ShowColumns
End Sub

Private Sub chkShowOrderInfo_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    If chkShowOrderInfo Then
        Me.ClubOrderSub.Height = boxSubFormNormal.Height - 250
        Me.ClubOrderSub.Top = boxSubFormNormal.Top
    Else
        Me.ClubOrderSub.Top = BoxBody.Top
        Me.ClubOrderSub.Height = BoxBody.Height - 250
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Check93_Click"
    Resume ExitSub
End Sub

Private Sub chkSize_Click()
    ShowColumns

End Sub

Private Sub ClubOrderSub_Exit(Cancel As Integer)
'    MsgBox "Exiting Sub Form"
    'Me.ClubOrderSub.Form.Recalc
    'don't dirty form unless you have to

'    MsgBox "Exiting Sub Form"

End Sub

Private Sub cmbOrderID_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    ' Find the record that matches the control.
    Dim rs As Object
    'SHOW ALL THE RECORDS
    Set rs = Me.Recordset.Clone

    rs.FindFirst "[ID] = " & Str(Me![cmbOrderID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbOrderID_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click

    'InvDollars = txtTotalPrice
    'DisDollars = txtTotalDiscount
    'InvQty = txtTotalQ

Exit_cmdClose_Click:
    DoCmd.Close
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub

Private Sub cmdDelOrder_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdDelOrder_Click

    If gbAllowDeleteDocument Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
        DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
        DoCmd.Close
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
    End If
Exit_cmdDelOrder_Click:
    Exit Sub

Err_cmdDelOrder_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelOrder_Click

End Sub


Private Sub cmdPreviousInvoiceReport_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPreviousInvoiceReport_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PrintInventoryReport"

    stLinkCriteria = "[ReportName]=" & "'26 Club Invoices'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , txtClubName



ExitSub:
    Exit Sub

Err_cmdPreviousInvoiceReport_Click:
    MsgBox Err.Description
    Resume ExitSub

End Sub

Private Sub cmdPrint_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrint_Click
    'added 6/5/2018
    InvDate = Now

    'added 02/16/2006
    'MsgBox "Reprice"
    RepriceClubOrder

    '20007-03-24
    'for whatever reason.. this (me.refresh) was causing a crash..
    'repair open data base error'
    'Office2003 sp2 fixed it MOST of the time...
    '
    'Me.Refresh

    If Not gcfHandleErrors Then Debug.Print Now

   ' MsgBox "calctotals"

    ClubOrderRecalcTotals (ID)
    If Me.Dirty Then
        'MsgBox "dirty"
        Me.Dirty = False
    End If





   ' MsgBox "AssignClubInvoiceNO"
    If IsNull(InvoiceNo) Or InvoiceNo = "" Then
        MsgBox "Assign Invoice Number!"
        AssignClubInvoiceNO CLng(ID)
        Me.Refresh
        If Not gcfHandleErrors Then Debug.Print "Invoice = " & InvoiceNo
       'cmbOrderID.Requery
       'Me.Refresh
    End If

    gstDocName = "ClubOrderPrintDialog"
    gstLinkCriteria = "[ID]=" & Me![ClubID]
    gstrInvoiceNo = InvoiceNo
   ' MsgBox "ClubOrderPrintDialog"
    DoCmd.OpenForm gstDocName, , , gstLinkCriteria, , , gstrInvoiceNo
    'DoCmd.Close acForm, Me.Name
Exit_cmdPrint_Click:
    Exit Sub

Err_cmdPrint_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrint_Click

End Sub


Private Sub cmdPurchasesByItem_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrevInvoices_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PrintInventoryReport"

    stLinkCriteria = "[ReportName]=" & "'24 Club History By Bottle By Month'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , txtClubName



ExitSub:
    Exit Sub

Err_cmdPrevInvoices_Click:
    MsgBox Err.Description
    Resume ExitSub
End Sub

Private Sub cmdRefresh_Click()
If gcfHandleErrors Then On Error GoTo Err_Command157_Click
    Me.ClubOrderSub.Form.Refresh

Exit_Command157_Click:
    Exit Sub

Err_Command157_Click:
    MsgBox Err.Description
    Resume Exit_Command157_Click


End Sub

Private Sub cmdResort_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdResort_Click


    Me.ClubOrderSub.Form.OrderByOn = False
    Me.ClubOrderSub.Form.OrderByOn = True

Exit_cmdResort_Click:
    Exit Sub

Err_cmdResort_Click:
    MsgBox Err.Description
    Resume Exit_cmdResort_Click

End Sub


Private Sub cmdSaveOrderGuide_Click()

Dim rstemp

'Dim rsOrder As New ADODB.Recordset
Dim rsOrder As Object
Set rsOrder = CreateObject("ADODB.Recordset")
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")
Dim comQDF1 As Object
Set comQDF1 = CreateObject("ADODB.COMMAND")
Dim comQDF2 As Object
Set comQDF2 = CreateObject("ADODB.COMMAND")


    Dim strSQL As String
    Dim strCnn As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    strSQL = "This will save the order guide in the order it is currently displayed." & vbCrLf & vbCrLf
    strSQL = strSQL & "Do you want to save the order guide now?"

    If MsgBox(strSQL, vbYesNo) = vbNo Then
        Exit Sub
    End If

    '1. open the grid as a recordset
    Set rstemp = Me.ClubOrderSub.Form.RecordsetClone


    '2. save the grid as a -1 order guide records for this club
    '   Note: an error in this routine indicates the previous save did NOT complete
    If gcfHandleErrors Then On Error GoTo ErrPreviousSave
    With rstemp
        If .RecordCount > 0 Then
            .MoveFirst
            rsOrder.Open "ClubOrderGuides", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            Do While Not .EOF
                rsOrder.AddNew
                rsOrder!ClubID = ClubID
                rsOrder!OrderID = -1
                rsOrder!LineID = .AbsolutePosition + 1
                rsOrder!LiquorItem = 255 ' 'True
                rsOrder!ItemID = !ItemID
                rsOrder!Par = !QuantityOrdered  'stuffed in here when creating order Par
                rsOrder!QuotedPrice = 0 '!QuotedPrice
                ' is this necessary? rsOrder!Description = !Description
                rsOrder.Update
                If Not gcfHandleErrors Then Debug.Print .AbsolutePosition; !LineID; !ItemID
                .MoveNext
            Loop
            rsOrder.Close
            .Close
        End If
    End With
    Set rstemp = Nothing
    Set rsOrder = Nothing

    If gcfHandleErrors Then On Error GoTo ErrRollBack
    'CurrentProject.Connection.BeginTrans
    '3. Delete original order guide records for this club
    strSQL = "DELETE  "
    strSQL = strSQL & "FROM ClubOrderGuides "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ClubID & ") "
    strSQL = strSQL & " AND ((ClubOrderGuides.OrderID)=1))"
   ' CurrentDb().Execute strSQL
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute , , adCmdText
    End With


    '4. change the -1 order guide records to 1 order guide
    strSQL = "UPDATE ClubOrderGuides "
    strSQL = strSQL & "SET ClubOrderGuides.OrderID = 1 "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ClubID & ") "
    strSQL = strSQL & "AND ((ClubOrderGuides.OrderID)=-1))"
    'CurrentDb().Execute strSQL
    With comQDF1
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute , , adCmdText
    End With

    '5. Delete any -1 order guide records for this club
    strSQL = "DELETE  "
    strSQL = strSQL & "FROM ClubOrderGuides "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ClubID & ") "
    strSQL = strSQL & " AND ((ClubOrderGuides.OrderID)=-1))"
    'CurrentDb().Execute strSQL

    With comQDF2
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute , , adCmdText
    End With
    'CurrentProject.Connection.CommitTrans
    'cnn1.Close
    'this forces the subform to resort

    optOrderBy = 1
    optOrderBy_Click

ExitSub:
    Set comQDF = Nothing
    Set comQDF1 = Nothing
    Set comQDF2 = Nothing
    'Set cnn1 = Nothing
    Exit Sub

ErrRollBack:
    strSQL = "There were errors saving the order guide for Club: " & ClubID
    MsgBox strSQL
    'cnn1.RollbackTrans
    'cnn1.Close
    Resume ExitSub

ErrLog:
    EMsgBox Me, "cmdSaveOrderGuide_Click"
    Resume ExitSub

ErrPreviousSave:
    strSQL = "There were errors saving the order guide for Club: " & ClubID
    MsgBox strSQL
    EMsgBox Me, "CommcmdSaveOrderGuide_Clickand116"

    Resume ExitSub

End Sub



Private Sub Form_Current()
Dim strSQL As String
Dim strSQLWhereDesc As String
Dim strSQLWhereCode As String
    Dim buildstr As String

If gcfHandleErrors Then On Error GoTo ErrLog
    cmbOrderID = ID
    With cmbOrderID
        txtClubName = .Column(1)
        txtClubContact = .Column(6)
        txtClubAddress1 = .Column(7)
        txtClubAddress2 = .Column(8)
        txtClubCityStZip = .Column(9)
        txtClubPhone = .Column(10)
        chkNonLiquor = .Column(11)
        txtDiscountRate = .Column(13)
        txtNetPrice = .Column(15)
        chkApplyDiscounts = .Column(16)
        chkParLevels = .Column(17)
        'added 2008-1-2
        'this works for mapquest
        'buildstr = "http://www.mapquest.com/maps/map.adp?formtype="
        'buildstr = buildstr & "&address=" & txtClubAddress1
        'buildstr = buildstr & "&city=" & Trim(.Column(19))
        'buildstr = buildstr & "&state=" & Trim(.Column(20))
        'buildstr = buildstr & "&zipcode=" & Trim(.Column(21))
        If Not IsNull(Trim(txtClubAddress1)) Or Trim(txtClubAddress1) <> "" Then
            buildstr = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q="
            buildstr = buildstr & Trim(txtClubAddress1) & ", " & Trim(.Column(19)) & ", " & Trim(.Column(20)) & " " & Trim(.Column(21))
            buildstr = buildstr & "&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
            buildstr = Replace(buildstr, " ", "+")
            lblMapQuest.Caption = "Click here for &Google Map to " & txtClubName
        Else
            lblMapQuest.Caption = "Click here for &Google Map"
            buildstr = "http://maps.google.com/maps"
        End If
        If .Column(22) <> "0" Then
            bFintechLiquorOnly = GetKMDSSettings("FintechLiquorOnly", "True")
            bFintech = True
            Label50.Caption = "Fintech"
            Label50.ForeColor = vbRed
        Else
            bFintech = False
            Label50.Caption = "Dispenser"
            Label50.ForeColor = vbBlue
        End If


    End With

    lblMapQuest.HyperlinkAddress = buildstr
    lblMapQuest.ControlTipText = buildstr

    'don't show cost on club orders!
    chkCost = chkNonLiquor
    chkExtCost = chkNonLiquor

'    If Nz(chkNonLiquor, 0) Then
'        'COMBINE LIQUOR AND RELATED 'CIGS AND
'        'strSQLWhereCode = "WHERE (((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID) <> 1) "
'        'strSQLWhereCode = strSQLWhereCode & " And ((Len([UPCItem].[ID])) > 3))"
'        'strSQLWhereDesc = "WHERE (((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID) <> 1)) "
'
'        strSQLWhereCode = "WHERE (Len([UPCItem].[ID]) > 3) "
'        strSQLWhereDesc = ""
'
'    Else
'        'strSQLWhereCode = "WHERE (((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID) = 1) "
'        'strSQLWhereCode = strSQLWhereCode & " And ((Len([UPCItem].[ID])) > 3))"
'        'strSQLWhereDesc = "WHERE (((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID) = 1)) "
'
'        strSQLWhereCode = "WHERE (Len([UPCItem].[ID]) > 3) "
'        strSQLWhereDesc = ""
'    End If
'
'    strSQL = "SELECT qryPurchaseOrderItemsWithTodaysPrice.ID, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Description, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.PkgSize, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Onhand, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Cost, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Sprice, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Star, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.ItemVendorID AS VID, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.UnitsPerCase, "
'    strSQL = strSQL & "nz([OnReserve],0) AS Reserved, "
'    strSQL = strSQL & "[Onhand]-nz([onReserve],0) AS Avail, "
'    strSQL = strSQL & "nz([OnOrder],0) AS Ordered, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID "
'    strSQL = strSQL & "FROM (qryPurchaseOrderItemsWithTodaysPrice "
'    strSQL = strSQL & "  LEFT JOIN qryItemOnOrder "
'    strSQL = strSQL & "   ON qryPurchaseOrderItemsWithTodaysPrice.ID = qryItemOnOrder.ItemID) "
'    strSQL = strSQL & "  LEFT JOIN qryItemOnReserve "
'    strSQL = strSQL & "   ON qryPurchaseOrderItemsWithTodaysPrice.ID = qryItemOnReserve.ItemID "
'    strSQL = strSQL & strSQLWhere
'    strSQL = strSQL & "ORDER BY qryPurchaseOrderItemsWithTodaysPrice.ID "


'    strSQL = "SELECT UPCItem.ID as CodeOrUPC, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Description, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.PkgSize, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Onhand, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Cost, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Sprice, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Star, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.ItemVendorID AS VID, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.UnitsPerCase, "
'    strSQL = strSQL & "nz([OnReserve],0) AS Reserved, "
'    strSQL = strSQL & "[Onhand]-nz([onReserve],0) AS Avail, "
'    strSQL = strSQL & "nz([OnOrder],0) AS Ordered, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID, "
'    strSQL = strSQL & "UPCItem.ItemID "
'    strSQL = strSQL & "FROM UPCItem INNER JOIN ((qryPurchaseOrderItemsWithTodaysPrice "
'    strSQL = strSQL & "              LEFT JOIN qryItemOnOrder ON qryPurchaseOrderItemsWithTodaysPrice.ID = qryItemOnOrder.ItemID) "
'    strSQL = strSQL & "              LEFT JOIN qryItemOnReserve ON qryPurchaseOrderItemsWithTodaysPrice.ID = qryItemOnReserve.ItemID) "
'    strSQL = strSQL & "             ON UPCItem.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID "


'    'You've been replaced by a parameter
'    'strSQL = strSQL & "WHERE (((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID) = 1))"
'    strSQL = strSQL & strSQLWhereCode
'    strSQL = strSQL & "ORDER BY Len([UPCITEM].[ID]), UPCItem.ID;"
'    'Me.ClubOrderSub.Form![ItemID].RowSource = strsql


'    strSQL = "SELECT Item.Description, "
'    strSQL = strSQL & "Item.ID, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.PkgSize, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Sprice, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.Star, "
'    strSQL = strSQL & "qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID "
'    strSQL = strSQL & "FROM qryPurchaseOrderItemsWithTodaysPrice "
'    strSQL = strSQL & " INNER JOIN Item "
'    strSQL = strSQL & "  ON qryPurchaseOrderItemsWithTodaysPrice.ID=Item.ID "
'    strSQL = strSQL & strSQLWhereDesc
'    strSQL = strSQL & "ORDER BY Item.Description"
'    'Me.ClubOrderSub.Form![Description].RowSource = strsql

    ShowColumns


    optOrderBy = 1
    optOrderBy_Click

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub



Private Sub Form_Delete(Cancel As Integer)
    If gbAllowDeleteDocument Then
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
        Cancel = True
    End If

End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Me.Form.InsideWidth = BoxBody.Width + 500
    Me.Form.InsideHeight = Me.FormHeader.Height + BoxBody.Height + Me.FormFooter.Height + 100

    cmbOrderID.BackColor = glnColorDropDown

    txtTotalQ.BackColor = glnColorTotals
    'txtTotalCost.BackColor = glnColorTotals
    txtTotalPrice.BackColor = glnColorTotals
    txtTotalDiscount.BackColor = glnColorTotals
    txtTotalNet.BackColor = glnColorTotals
    txtOrdDollars.BackColor = glnColorTotals
    txtInvQty.BackColor = glnColorTotals

    txtClubName.BackColor = glnColorInfo
    txtClubContact.BackColor = glnColorInfo
    txtClubAddress1.BackColor = glnColorInfo
    txtClubAddress2.BackColor = glnColorInfo
    txtClubCityStZip.BackColor = glnColorInfo
    txtClubPhone.BackColor = glnColorInfo
    txtLiquorDollars.BackColor = glnColorInfo
    txtLiquorDiscount.BackColor = glnColorInfo
    txtNonLiquorDollars.BackColor = glnColorInfo
    txtNonLiquorDiscount.BackColor = glnColorInfo

    Order.BackColor = glnColorInfo
    OrdDate.BackColor = glnColorInfo
    InvoiceNo.BackColor = glnColorInfo
    InvDate.BackColor = glnColorInfo
    cmbClubNotes.BackColor = glnColorInfo



ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub



Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
    SetFormSize Me

End Sub


Private Sub optOrderBy_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    Me.ClubOrderSub.Form.OrderByOn = False
    Select Case optOrderBy
    Case 1 'lineid
        Me.ClubOrderSub.Form.OrderBy = "LineID"

    Case 2 ' item id
        Me.ClubOrderSub.Form.OrderBy = "ItemID"
    Case 3 ' desc
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Description.Description,itemid"
    Case 4 ' quantity
        Me.ClubOrderSub.Form.OrderBy = "QuantityShipped, ItemID"
    Case 5 ' price
        Me.ClubOrderSub.Form.OrderBy = "Price, ItemID"
    Case 6 ' cost
        Me.ClubOrderSub.Form.OrderBy = "OrderCost, ItemID"
    Case 7 ' ext price
        Me.ClubOrderSub.Form.OrderBy = "ExtPrice, ItemID"
    Case 8 ' ext cost
        Me.ClubOrderSub.Form.OrderBy = "ExtCost, ItemID"
    Case 9 ' size; code
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Size.PkgSize, ItemID"
    Case 10 ' size; desc
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Size.PkgSize, Lookup_Description.Description"
    Case 11 'category; code
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Category.Descr, ItemID"
    Case 12 'category; desc
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Category.Descr, Lookup_Description.Description"
    Case 13 'category; price
        Me.ClubOrderSub.Form.OrderBy = "Lookup_Category.Descr, Price"
   End Select
    Me.ClubOrderSub.Form.OrderByOn = True
ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 3709
        'SEARCH KEY NOT FOUND... OrderByOn gives an error message when there are
        'no lines!

    Case Else
        MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "optOrderBy_Click"
    End Select

    Resume ExitSub
End Sub



Private Sub ShowColumns()
If gcfHandleErrors Then On Error GoTo ErrLog

        'Me.PurchaseOrdersOnFileSub.Form![ExtCost].ColumnHidden = Not chkDisplayCost
    Me.ClubOrderSub.Form![LineID].ColumnHidden = Not chkLineID
    Me.ClubOrderSub.Form![ItemID].ColumnHidden = Not chkItemID
    Me.ClubOrderSub.Form![Description].ColumnHidden = Not chkDesc
    Me.ClubOrderSub.Form![Quantity].ColumnHidden = Not chkQuantity
    Me.ClubOrderSub.Form![Price].ColumnHidden = Not chkPrice
    Me.ClubOrderSub.Form![Cost].ColumnHidden = Not chkCost
    Me.ClubOrderSub.Form![ExtPrice].ColumnHidden = Not chkExtPrice
    Me.ClubOrderSub.Form![ExtCost].ColumnHidden = Not chkExtCost
    Me.ClubOrderSub.Form![Size].ColumnHidden = Not chkSize
    Me.ClubOrderSub.Form![Category].ColumnHidden = Not chkCategory



ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ShowColumns"
    Resume ExitSub
End Sub
