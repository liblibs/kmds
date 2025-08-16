Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nMinOrderQuantity As Integer
Dim strStockNumber As String

Private Sub chkShowExtended_AfterUpdate()
    PriceGroupID.Visible = chkShowExtended
    VendorID.Visible = chkShowExtended
    Cost.Visible = chkShowExtended
    Deposit.Visible = chkShowExtended
    QtyPack.Visible = chkShowExtended
    QtyPackPrice.Visible = chkShowExtended
    StockNumber.Visible = chkShowExtended
    UnitsPerCase.Visible = chkShowExtended
    If chkShowExtended Then
        SetKMDSSettings "ItemsSetupShowExtendedFields", "True"
    Else
        SetKMDSSettings "ItemsSetupShowExtendedFields", "False"
    End If
End Sub


Private Sub cmbSizeID_Enter()
    If IsNull(cmbSizeID) Then
        cmbSizeID = 0
    End If
    If cmbSizeID = "" Then
        cmbSizeID = 0
    End If
    'If not gcfHandleErrors then Debug.Print "sizeid"; cmbSizeID
End Sub

Private Sub cmdDupe_Click()
    Dim rs As Object
    Dim lngID As Long
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = CurrentDb().OpenRecordset("select * from Item Where ID =" & CStr([Forms]![Item]![ID]))
    If rs.RecordCount > 0 Then
        rs.MoveLast
        If MsgBox("Duplicate " & rs!ID & " " & rs!Description & "?", vbYesNo) = vbYes Then
            SalesDepartmentID = rs!SalesDepartmentID
            Description = rs!Description
            Price = rs!Price
            cmbSizeID = rs!SizeID
            PriceGroupID = rs!PriceGroupID
            VendorID = rs!VendorID
            Cost = rs!Cost
            Deposit = rs!Deposit
            QtyPack = rs!QtyPack
            QtyPackPrice = rs!QtyPackPrice
            StockNumber = rs!StockNumber
            UnitsPerCase = rs!UnitsPerCase


            SalesDepartmentID_Exit (False)
            chkShowExtended = True
            chkShowExtended_AfterUpdate

            Description.SetFocus
        End If
    End If

    rs.Close
    Set rs = Nothing
ExitSub:
    Exit Sub

ErrLog:
    SalesDepartmentID.SetFocus
    Select Case Err.Number
    Case 2450
        MsgBox "D-1 Item Screen must be running!"
    Case Else
        EMsgBox Me, "cmdDupe_Click"
    End Select

    Resume ExitSub

End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub
End Sub



Private Sub Form_AfterInsert()
    gbRunMailChanges = True
    DoCmd.Close

End Sub

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True
    DoCmd.Close

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    If IsNull(SalesDepartmentID) Then
        SalesDepartmentID.SetFocus
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_BeforeInsert"
    Resume ExitSub
End Sub



Private Sub Form_Current()
'    chkShowExtended = False
    chkShowExtended_AfterUpdate

End Sub


Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Dim rs As Object
    nMinOrderQuantity = 1

    If Me.OpenArgs <> "" Then
        'setting up a new item from the po screen

        SalesDepartmentID = 1

        ID = Me.OpenArgs
        cmbSizeID = Right(ID, 2)
        Description.SetFocus
        'added 2008-1-2 ... look it up in OLCCWebItems
        'changed 2010-10-12 ... use info from olcc ftp files
        Set rs = CurrentDb().OpenRecordset("select * FROM dbo_OLCCItem where CODE =" & ID, dbOpenForwardOnly, dbSeeChanges)
        With rs
            If Not rs.EOF Then
                SalesDepartmentDesc = SalesDepartmentID

                Description = !Description
                nMinOrderQuantity = !UnitsPerCase
                UnitsPerCase = !UnitsPerCase
                Price = !Price
                strStockNumber = !StockStatus
                VendorID = 1
                MsgBox "PLEASE VERIFY PRICE" & vbCrLf & vbCrLf & "Price shown is as of " & Format(!SetupDate, "mm/dd/yyyy"), vbCritical, "VERIFY PRICE"
            End If
           .Close
        End With
        Set rs = Nothing


    Else
        'added 2006-10-13
        cmbSizeID = 0 ' default to not assigned
        SalesDepartmentID.SetFocus
    End If
    If GetKMDSSettings("ItemsSetupShowExtendedFields", "False") = "False" Then
        chkShowExtended.Value = False
    Else
        chkShowExtended.Value = True
    End If
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Form Load"
    Resume ExitSub

End Sub

Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
    SetFormSize Me


End Sub


Private Sub ID_Exit(Cancel As Integer)
    Dim rs As Object
    Dim lngID As Long
    Dim dblID As Double



    If gcfHandleErrors Then On Error GoTo ErrLog
    nMinOrderQuantity = 1
    'ADDED 31416

    If Not IsNumeric(ID) Then
        If IsNull(ID) Or ID = "" Then
        Else
            MsgBox "This field must be a number!"
            Cancel = True
            Exit Sub
        End If
    Else
        Set rs = CurrentDb().OpenRecordset("select Description,Price from Item Where ID =" & ID)
        If rs.RecordCount > 0 Then
            MsgBox "This code is setup up for " & rs!Description & " " & rs!Price
            Cancel = True
            ID = ""
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing



        If SalesDepartmentID <> 1 Then
            If ID > 99999999 Then ' 10,000,000
                MsgBox "This Non-Liquor item can NOT be GREATER than 7 digits, please assign a number less than 9999999"
                Cancel = True
                Exit Sub
            End If

        Else
            If ID < 10000000000# Or ID > 99999999999# Then
                MsgBox "This Liquor item must be in the 999-999999-99 range"
                Cancel = True
                Exit Sub
            Else
                cmbSizeID = Right(ID, 2)
                'added 2008-01-02
                'changed 2010-10-12 ... use info from olcc ftp files
                Set rs = CurrentDb().OpenRecordset("select * FROM dbo_OLCCItem where CODE =" & ID, dbOpenForwardOnly, dbSeeChanges)
                With rs
                    If Not rs.EOF Then
                        Description = !Description
                        nMinOrderQuantity = !UnitsPerCase
                        UnitsPerCase = !UnitsPerCase
                        Price = !Price
                        strStockNumber = !StockStatus
                        MsgBox "PLEASE VERIFY PRICE" & vbCrLf & vbCrLf & "Price shown is as of " & Format(!SetupDate, "mm/dd/yyyy"), vbCritical, "VERIFY PRICE"
                    End If
                   .Close
                End With
                Set rs = Nothing

            End If
        End If
        'END 31416

    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ID_Exit"
    Resume ExitSub
End Sub


Private Sub SalesDepartmentDesc_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    SalesDepartmentID = SalesDepartmentDesc
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "SalesDepartmentDesc_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub SalesDepartmentDesc_Enter()
'    SendKeys "%{DOWN}"
End Sub


Private Sub SalesDepartmentID_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    SalesDepartmentDesc = SalesDepartmentID
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "SalesDepartmentID_AfterUpdate"
    Resume ExitSub
End Sub




Private Sub SalesDepartmentID_Exit(Cancel As Integer)
    Dim rs As Object
    Dim lngID As Long

    If gcfHandleErrors Then On Error GoTo ErrLog
    If SalesDepartmentID = 1 Then
        ID.SetFocus
        VendorID = 1
    Else
        'ADDED 31416
        'Set RS = CurrentDb().OpenRecordset("select max(ID) as MaxID from Item")
        Set rs = CurrentDb().OpenRecordset("select max(ID) as MaxID from Item WHERE ID < 999999999") '10,000,000

        rs.MoveLast
        ID = rs!MaxID + 1
        'If ID < 1000000 Then ID = 1000000 '1,000,000

        'txtUPC.SetFocus
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "SalesDepartmentID_Exit"
    Resume ExitSub
End Sub

Private Sub txtUPC_Exit(Cancel As Integer)
    Dim rs As Object

    Dim lngID As Long

If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNull(txtUPC) Or txtUPC = 0 Or txtUPC = "" Then
        txtUPC = ID
    End If

    Set rs = CurrentDb().OpenRecordset("select * from UPCItem Where ID ='" & CStr(txtUPC) & "'")
    If rs.RecordCount > 0 Then
        MsgBox "This code has been used!" & vbCrLf & vbCrLf & "It is in use for item #" & rs!ItemID
        Cancel = True
        txtUPC = ""
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "txtUPC_Exit"
    Resume ExitSub
End Sub

Private Sub cmdSave_Click()


    Dim strSQL As String
    Dim strSize As String
    Dim rs As Object
    Dim nSize As Integer
    Dim nVendorID As Integer


    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNumeric(SalesDepartmentID) Then
        MsgBox ("SalesDepartment ID must have a value!")
        SalesDepartmentID.SetFocus
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

    If IsNull(Description) Or Description = "" Then
        MsgBox ("Description can not be blank!")
        Description.SetFocus
        Exit Sub
    End If

    If IsNull(Price) Or Price = "" Then
        Price = 0
    End If

    If MsgBox("Save This Item?", vbYesNo, "OK to save?") = vbYes Then
        gbRunMailChanges = True
        'if this is a liquor item then try to determine size and set those defaults
        If SalesDepartmentID = 1 Then
            nSize = Right(ID, 2)
            strSQL = ""
            strSize = ""
            Select Case nSize
            Case 5
                strSize = "50ML"
            Case 10
                strSize = "1LTR"
            Case 17
                strSize = "1.75L"
            Case 20
                strSize = "200ML"
            Case 37
                strSize = "375ML"
            Case 75
                strSize = "750ML"
            Case 70
                strSize = "700ML"
            Case 50
                strSize = "500ML"

            Case Else 'size not determined
                MsgBox "Liquor items must end with 05,10,17,20,50,37,70,75", vbCritical, "Invalid ID number (Wrong size)"
                ID.SetFocus
                Exit Sub
            End Select
            If strSize <> "" Then
                nVendorID = 1
                If strStockNumber = "S" Then 'special order sets vendor to 0
                    strSQL = "Special Order items MUST be ordered through OLAS" & vbCrLf
                    strSQL = strSQL & "Are you sure you want to add " & Description & "?"
                    If MsgBox(strSQL, vbYesNo) <> vbYes Then
                        GoTo NoSpecialOrder
                    End If
                    nVendorID = 0
                End If

                strSQL = "SELECT Description, "
                strSQL = strSQL & "Oz , "
                strSQL = strSQL & "CaseQty, "
                strSQL = strSQL & "ID "
                strSQL = strSQL & "FROM [Size] "
                strSQL = strSQL & "WHERE (((Description)='" & strSize & "'))"
                Set rs = CurrentDb().OpenRecordset(strSQL)
                If rs.RecordCount > 0 Then
                    If IsNull(UnitsPerCase) Then UnitsPerCase = 1
                    If IsNull(nMinOrderQuantity) Then nMinOrderQuantity = 1
                    If nMinOrderQuantity = 0 Then nMinOrderQuantity = 1
                    If UnitsPerCase = 0 Then UnitsPerCase = 1
                    'ADDED 01/13/2013
                    If UnitsPerCase = "" Then UnitsPerCase = 1
                    Description = Replace(Description, "'", "")
                    Description = Replace(Description, Chr(34), "")

                    strSQL = "INSERT INTO Item ( ID, "
                    strSQL = strSQL & "Description, "
                    strSQL = strSQL & "SalesDepartmentID, "

                    strSQL = strSQL & "SizeID, "
                    strSQL = strSQL & "UnitsPerCase, "
                    strSQL = strSQL & "UnitsPerOrder, "
                    strSQL = strSQL & "VendorID, "
                    strSQL = strSQL & "StockNumber, "
                    strSQL = strSQL & "Cost, "
                    strSQL = strSQL & "Price ) "

                    strSQL = strSQL & "SELECT " & ID & " AS ID, "
                    strSQL = strSQL & "'" & Description & "' AS Description, "
                    strSQL = strSQL & "1 AS SalesDepartmentID, "
                    strSQL = strSQL & Nz(rs!ID, 0) & " AS SizeID, "
                    strSQL = strSQL & Nz(UnitsPerCase, 0) & " AS UnitsPerCase, "
                    strSQL = strSQL & Nz(nMinOrderQuantity) & " AS UnitsPerOrder, "
                    strSQL = strSQL & nVendorID & " AS VendorID, "
                    strSQL = strSQL & "'" & strStockNumber & "' AS StockNumber, "
                    strSQL = strSQL & Price & " AS Cost,"
                    strSQL = strSQL & Price & " AS Price"


                Else
                    strSQL = ""
                End If 'recordcount = 0
                rs.Close
            Else
                MsgBox "A valid SIZE is required", vbCritical, "Item not saved"
                cmbSizeID.SetFocus
                Exit Sub
            End If 'strsize = ""
        End If 'SalesDepartmentID = 1

        If strSQL = "" Then
            If cmbSizeID = "" Then cmbSizeID = 0
            If Cost = "" Then Cost = 0
            If Deposit = "" Then Deposit = 0
            If QtyPack = "" Then QtyPack = 0
            If QtyPackPrice = "" Then QtyPackPrice = 0
            If UnitsPerCase = "" Then UnitsPerCase = 0
            If VendorID = "" Then VendorID = 0
            If PriceGroupID = "" Then PriceGroupID = 0
            If Price = "" Then Price = 0

            strSQL = "INSERT INTO Item ( ID, "
            strSQL = strSQL & "Description, "
            strSQL = strSQL & "SalesDepartmentID, "
            strSQL = strSQL & "SizeID, "

            strSQL = strSQL & "Cost, "
            strSQL = strSQL & "Deposit, "
            strSQL = strSQL & "QtyPack, "
            strSQL = strSQL & "QtyPackPrice, "
            strSQL = strSQL & "StockNumber, "
            strSQL = strSQL & "UnitsPerCase, "
            strSQL = strSQL & "VendorID, "
            strSQL = strSQL & "PriceGroupID, "

            strSQL = strSQL & "Price ) "
            strSQL = strSQL & "SELECT " & ID & " AS ID, "
            strSQL = strSQL & "'" & Description & "' AS Description, "
            strSQL = strSQL & SalesDepartmentID & " AS SalesDepartmentID, "
            strSQL = strSQL & cmbSizeID & " AS SizeID, "


            strSQL = strSQL & Nz(Cost, 0) & ", "
            strSQL = strSQL & Nz(Deposit, 0) & ", "
            strSQL = strSQL & Nz(QtyPack, 0) & ", "
            strSQL = strSQL & Nz(QtyPackPrice, 0) & ", "
            strSQL = strSQL & "'" & Nz(StockNumber, "") & "', "
            strSQL = strSQL & Nz(UnitsPerCase, 0) & ", "
            strSQL = strSQL & Nz(VendorID, 0) & ", "
            strSQL = strSQL & Nz(PriceGroupID, 0) & ", "

            strSQL = strSQL & Nz(Price, 0) & " AS Price"
        End If
        CurrentDb().Execute (strSQL)

        strSQL = "INSERT INTO UPCItem ( ID, ItemID ) "
        strSQL = strSQL & "SELECT " & ID & " AS ID, "
        strSQL = strSQL & ID & " AS ItemID"
        If gcfHandleErrors Then On Error Resume Next
        CurrentDb().Execute (strSQL)

        If txtUPC > "" And txtUPC <> ID Then
            strSQL = "INSERT INTO UPCItem ( ID, ItemID ) "
            strSQL = strSQL & "SELECT '" & txtUPC & "' AS ID, "
            strSQL = strSQL & ID & " AS ItemID"
            If gcfHandleErrors Then On Error Resume Next
            CurrentDb().Execute (strSQL)
        End If
NoSpecialOrder:
        SalesDepartmentID = ""
        SalesDepartmentDesc = ""
        ID = ""
        txtUPC = ""
        Description = ""
        cmbSizeID = ""
        Price = ""

        Cost = ""
        QtyPack = ""
        QtyPackPrice = ""
        StockNumber = ""
        Deposit = ""
        UnitsPerCase = ""
        VendorID = ""
        PriceGroupID = ""
        'let user decide if it is shown
        'chkShowExtended = False
        'chkShowExtended_AfterUpdate
        SalesDepartmentID.SetFocus



        Me.Visible = False
        'DoCmd.Close acForm, Me.Name

    Else
        'Cancel = True
        Exit Sub
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Price_Exit"
    Resume ExitSub

End Sub
