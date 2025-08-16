Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'simple down & dirty inventory input screen designed to use a bluetooth scanner and allow
'user to run program 'from a distance'  (Keep keyboard entries to a bare minimum)
'input card with quantities will be provided so the count can be entered with a scanner

'mrsScanBatch is the lookup query based on vwUPCItemsWithCurrentPricing
'mrsScanBatchLineCounts is a table that holds the counts

Dim mrsScanBatch As Object
'Dim mrsScanBatchLineCounts As New ADODB.Recordset
Dim mrsScanBatchLineCounts As Object

Dim mrsSBLC As Recordset
Dim mrsItemTotals As Recordset


Private Sub Form_Load()
'    Application.SetOption "Error Trapping", 1
    Dim strADD As String
    'Dim rsOrder As New ADODB.Recordset
    If gcfHandleErrors Then On Error GoTo ErrLog
    'makesure the list box showing previous scans is emptied
    SetupScreens
    StatusLine "Opening Tables", vbYellow

    Set mrsScanBatch = CurrentDb().OpenRecordset("qryScanBatch")
    txtOnShelfDate = Now ' this is when the above recordset was created

    Set mrsScanBatchLineCounts = CreateObject("ADODB.Recordset")
    mrsScanBatchLineCounts.Open "ScanBatchLineCounts", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    lstLocation.Visible = True
    lstLocation.SetFocus
ExitSub:
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  Form_Load", vbRed
    Resume ExitSub
End Sub

'Private Sub lstLocation_DblClick(Cancel As Integer)
'    StatusLine "Location: " & lstLocation, vbRed
'    Me.Caption = "Scan Batch at " & lstLocation
'    txtLocation = lstLocation
'    txtCode.SetFocus
'End Sub

Private Sub lstLocation_Enter()
    StatusLine "Please select the location and Press Enter", vbYellow
End Sub

Private Sub lstLocation_Exit(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    StatusLine "Location: " & lstLocation, vbRed
    Me.Caption = "Scan Batch at " & lstLocation
    txtLocation = lstLocation
    If lstLocation = "2. Backroom" Then
    Else

    End If

    txtCode.SetFocus
ExitSub:
    Exit Sub
ErrLog:
    StatusLine "Error: " & Err.Number & " " & Err.Description & "  lstLocation_Exit", vbRed
    Resume ExitSub

End Sub

'Private Sub ScanBatchLineCountsSub_Exit(Cancel As Integer)
'    lstEntries.Visible = True
'End Sub


Private Sub txtCode_DblClick(Cancel As Integer)
    DoCmd.OpenForm "ScanBatch1"
End Sub

Private Sub txtCode_Enter()
    lstLocation.Visible = False
    txtCode = ""
    StatusLine "Ready for next Item. Scan on! ", vbYellow
End Sub

Private Sub txtCode_Exit(Cancel As Integer)
   ' Find the record that matches the control.
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNull(txtCode) Then txtCode = CheckScanSheet(txtCode)
    If Len(txtCode) = 0 Then
        Cancel = True
        Beep
        StatusLine "Scan Code can not be empty! (Press 99998 to bypass)", vbRed
        Exit Sub
    End If

    If Len(txtCode) < 4 Then 'the user wants to add quantity to last count
        Cancel = True
        If Not IsNumeric(txtCode) Then 'quantity has to be a number
            Beep
            StatusLine "Adding quantity to last scanned item.. must be numeric", vbRed
            BlankInfoTextBoxes
            Exit Sub
        End If

        If lstEntries.ListCount = 1 Then ' there are no items in the list
            Beep
            StatusLine "No Last Scanned Item", vbRed
            BlankInfoTextBoxes
            Exit Sub
        End If
        If CLng(txtCode) > 99 Then
            StatusLine "Count must less than 100!", vbRed
            Beep
            BlankInfoTextBoxes
            Exit Sub
        End If

        'this is the bottom of the list..
'        txtCode = lstEntries.Column(0, lstEntries.ListCount - 1)
'        txtItemID = lstEntries.Column(1, lstEntries.ListCount - 1)
'        txtDescription = lstEntries.Column(2, lstEntries.ListCount - 1)
'        txtOnShelf = lstEntries.Column(4, lstEntries.ListCount - 1)

        'this is the top of the list
        txtCount = CInt(txtCode)
        txtCode = lstEntries.Column(0, 1)
        txtItemID = lstEntries.Column(1, 1)
        txtDescription = lstEntries.Column(2, 1)
        txtPrice = lstEntries.Column(3, 1)
        txtOnShelf = lstEntries.Column(5, 1)
        txtInputMethod = "Add quantity"
        AddEntryToListBox
        StatusLine "Ready for next Item", vbYellow
        Exit Sub
    End If

    Select Case txtCode
    Case 99999 '- User wants to remove last Item
        If lstEntries.ListCount > 1 Then
        'this is the bottom of the list..
'            txtCode = lstEntries.Column(0, lstEntries.ListCount - 1)
'            txtItemID = lstEntries.Column(1, lstEntries.ListCount - 1)
'            txtDescription = lstEntries.Column(2, lstEntries.ListCount - 1)
'            txtCount = 0 - lstEntries.Column(3, lstEntries.ListCount - 1)
'            txtOnShelf = lstEntries.Column(4, lstEntries.ListCount - 1)

            txtCode = lstEntries.Column(0, 1)
            txtItemID = lstEntries.Column(1, 1)
            txtDescription = lstEntries.Column(2, 1)
            txtPrice = lstEntries.Column(3, 1)
            txtCount = 0 - lstEntries.Column(4, 1)
            txtOnShelf = lstEntries.Column(5, 1)
            txtInputMethod = "Remove last"
            AddEntryToListBox
            lstEntries.RemoveItem (1)
            lstEntries.RemoveItem (1)
            StatusLine "Removing Last item!", vbYellow
        Else
            StatusLine "There are no more items to remove!", vbRed
            Beep
            BlankInfoTextBoxes
        End If
        Cancel = True
    Case 99998 ' allow me to click somewhere
        txtDescription.SetFocus
    Case Else
        With mrsScanBatch
            .FindFirst "[UPCCode] = '" & Me![txtCode] & "'"
             If Not .NoMatch Then
                txtItemID = !ID
                txtDescription = !Description
                txtPrice = Format(!Price, "#,##0.00")
                txtOnShelf = !OnShelf
                txtInputMethod = "Normal entry"
                txtCount.SetFocus
            Else
                StatusLine "Item not found", vbRed
                Beep
                Sleep (250)
                Beep
                Sleep (250)
                Beep
                Cancel = True
                BlankInfoTextBoxes
                'txtCode = ""
                'txtDescription = ""
                'txtPrice = ""
            End If
        End With
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "txtCode_Exit"
    StatusLine Err.Number & " " & Err.Description & " txtCode_Exit", vbRed
    Resume ExitSub
End Sub


Private Sub txtCount_Enter()
    StatusLine "Please enter count for " & txtDescription, vbYellow
End Sub

Private Sub txtCount_Exit(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog

    'is this off of the scan sheet?
    If Not IsNull(txtCount) Then txtCount = CheckScanSheet(txtCount)
    If Not IsNumeric(txtCount) Then
        Beep
        StatusLine "Count must be a number!", vbRed
        txtCount = ""
        Cancel = True
        Exit Sub
    End If

    '- User wants to remove last Item
    If txtCount = 99999 And lstEntries.ListCount > 1 Then
        txtCode = lstEntries.Column(0, 1)
        txtItemID = lstEntries.Column(1, 1)
        txtDescription = lstEntries.Column(2, 1)
        txtCount = 0 - lstEntries.Column(3, 1)
        txtOnShelf = lstEntries.Column(4, 1)
        txtInputMethod = "Remove last entry"
        AddEntryToListBox
        lstEntries.RemoveItem (1)
        lstEntries.RemoveItem (1)
        StatusLine "Removing Last item!", vbYellow
        Exit Sub
    End If

    If CLng(txtCount) > 99 Then
        StatusLine "Count must less than 100!", vbRed
        Beep
        txtCount = ""
        Cancel = True
        Exit Sub
    End If

    If CLng(txtCount) = 0 Then 'don't add a '0' count
        txtCode = ""
        txtDescription = ""
        txtPrice = ""
        txtCount = ""
        Exit Sub
    End If

    StatusLine "Adding scanned item to counts", vbYellow
    AddEntryToListBox

ExitSub:

    Exit Sub
ErrLog:

    StatusLine Err.Number & " " & Err.Description & " txtCount_Exit", vbRed
    Resume ExitSub
End Sub

Private Sub AddEntryToListBox()

    Dim strADD As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    StatusLine "Adding to list", vbYellow

    If AddLineCounts Then
        ShowLastScanInfo
        strADD = txtCode & ";"
        strADD = strADD & txtItemID & ";"
        strADD = strADD & txtDescription & ";"
        strADD = strADD & txtPrice & ";"
        strADD = strADD & txtCount & ";"
        strADD = strADD & txtOnShelf & ";"
        strADD = strADD & txtLocation & ";"
        strADD = strADD & txtInputMethod & ";"
        lstEntries.AddItem (strADD), 1
        txtCode = ""
        txtDescription = ""
        txtPrice = ""
        txtCount = ""
    End If
ExitSub:

    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "AddEntryToListBox"
    StatusLine Err.Number & " " & Err.Description & " AddEntryToListBox", vbRed
    Resume ExitSub
End Sub

Private Sub SetupScreens()
    Dim nI As Integer
    Dim strADD As String
    If gcfHandleErrors Then On Error GoTo ErrLog

    'first populate the list box with previous entries
    For nI = 0 To lstEntries.ListCount - 1
        lstEntries.RemoveItem (0)
    Next
    strADD = "UPC Code" & ";"
    strADD = strADD & "Item Id" & ";"
    strADD = strADD & "Description" & ";"
    strADD = strADD & "Price" & ";"
    strADD = strADD & "Count" & ";"
    strADD = strADD & "Shelf" & ";"
    strADD = strADD & "Location" & ";"
    strADD = strADD & "InputMethod" & ";"
    lstEntries.AddItem (strADD), 0
    'cmbEntries.AddItem (strAdd), 0

    Set mrsSBLC = Me.RecordsetClone
    With mrsSBLC
        If .RecordCount > 0 Then
            .MoveFirst
            txtFirstScanDate = !SetupDate
            .MoveLast
            txtLastScanDate = !SetupDate
            txtScanCounts = .RecordCount
            .MoveFirst
            Do While Not .EOF
                strADD = !ScanCode & ";"
                strADD = strADD & !ItemID & ";"
                strADD = strADD & !Description & ";"
                strADD = strADD & Format(!Price, "#,##0.00") & ";"
                strADD = strADD & !CountQuantity & ";"
                strADD = strADD & !OnShelf & ";"
                strADD = strADD & !Location & ";"
                strADD = strADD & !InputMethod & ";"
                lstEntries.AddItem (strADD), 1
                'cmbEntries.AddItem (strAdd), 1
                .MoveNext
            Loop
        End If
    End With

    '2. format the hidden form so it is exactly beneath list box
    With lstEntries
'        ScanBatchLineCountsSub.Top = .Top
'        ScanBatchLineCountsSub.Left = .Left
'        ScanBatchLineCountsSub.Height = .Height
'        ScanBatchLineCountsSub.Width = .Width
    End With

    '3. hide and seek with the 2 controls
    lstEntries.Visible = True
'    ScanBatchLineCountsSub.Visible = False
'    ScanBatchLineCountsSubLabel.Visible = False

    '4. locations are stored in the settings tables
    For nI = 0 To lstLocation.ListCount - 1
        lstLocation.RemoveItem (0)
    Next

    'up to 8 locations.. first 4 are provided
    lstLocation.AddItem "Select a Location  (Add More Locations in G-5 -> Settings)"
    strADD = GetKMDSSettings("ScanBatchLoc1", "1. Storeshelf")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc2", "2. Backroom")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc3", "3. Overflow")
    lstLocation.AddItem strADD
    strADD = GetKMDSSettings("ScanBatchLoc4", "4. Display")
    lstLocation.AddItem strADD
'    strAdd = GetKMDSSettings("ScanBatchLoc5", "5. Not Assigned")
'    If strAdd <> "5. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc6", "6. Not Assigned")
'    If strAdd <> "6. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc7", "7. Not Assigned")
'    If strAdd <> "7. Not Assigned" Then lstLocation.AddItem strAdd
'    strAdd = GetKMDSSettings("ScanBatchLoc8", "8. Not Assigned")
'    If strAdd <> "8. Not Assigned" Then lstLocation.AddItem strAdd

    '6. Set the screen size
    Me.InsideHeight = Box67.Height + FormFooter.Height + 200
    Me.InsideWidth = Box67.Width + 200

    '7 the location list box ???
    lstLocation.Top = txtCode.Top
    lstLocation.Width = txtDescription.Width
    lstLocation.Height = txtCode.Height * 4
ExitSub:
    Exit Sub
ErrLog:
    StatusLine Err.Number & " " & Err.Description & " ClearLog", vbRed
    Resume ExitSub
End Sub

Private Function AddLineCounts() As Boolean
    If gcfHandleErrors Then On Error GoTo ErrLog
    AddLineCounts = True
    With mrsScanBatchLineCounts
         .AddNew
         !ScanBatchHeaderID = glngScanBatchHeaderID
         !ItemID = txtItemID
         !CashierID = txtCashierID
         !InputMethod = txtInputMethod
         !Location = txtLocation
         !ScanCode = txtCode
         !CountQuantity = txtCount
         !OnShelf = txtOnShelf
         !OnShelfDate = txtOnShelfDate

         .Update
    End With
ExitSub:
    Exit Function
ErrLog:
    AddLineCounts = False
    Beep
    Sleep 250
    Beep
    Sleep 250
    Beep
    Sleep 250

    StatusLine "SCAN NOT SAVED! " & Err.Number & " " & Err.Description & " AddLineCounts", vbRed
    MsgBox "SCAN NOT SAVED! " & Err.Number & " " & Err.Description, , "AddLineCounts"
    Resume ExitSub
End Function

Private Sub StatusLine(message As String, Color As Long)
    txtStatus = message
    txtStatus.BackColor = Color
End Sub

Private Function CheckScanSheet(InputString As String) As String
    Dim intItemId As Long
    CheckScanSheet = InputString
    'copyright KMDS :)
    If Len(InputString) = 12 And Left(InputString, 6) = "000000" Then
        CheckScanSheet = CStr(CLng(Mid(InputString, 7, 5)))
    End If
    'check for olcc label 001402
    If Len(InputString) = 6 And Left(InputString, 1) = "0" And IsNumeric(InputString) Then
        intItemId = CLng(InputString) 'effectively trim leading zeros from inputstring
        CheckScanSheet = CStr(intItemId)
    End If

End Function

Private Sub ShowLastScanInfo():

    Dim strSQL As String
    Dim rs As Recordset
    txtItemIDLast = txtItemID
    txtCodeLast = txtCode
    txtDescriptionLast = txtDescription
    txtLastScanDate = Now
    With mrsSBLC
        .Requery
        .MoveLast
        txtScanCounts = .RecordCount
    End With

    strSQL = "SELECT ItemID, Sum(CountQuantity) AS CountQ, Last(OnShelf) AS OS, Count(CountQuantity) AS ScanQ "
    strSQL = strSQL & "From ScanBatchLineCounts GROUP BY ScanBatchLineCounts.ItemID "
    strSQL = strSQL & "HAVING (ItemID= " & CLng(txtItemID) & " )"
'    Set mrsScanBatch = CurrentDb().OpenRecordset("qryScanBatch")
    Set rs = CurrentDb().OpenRecordset(strSQL)

    With rs
        If .RecordCount Then
            .MoveFirst
            txtCountLast = !Countq
            txtOnShelfLast = !OS
            txtScanQ = !ScanQ
            txtVar = !Countq - !OS
            If txtVar = 0 Then txtVar.BackColor = vbWhite
            If txtVar < 0 Then txtVar.BackColor = vbRed
            If txtVar > 0 Then txtVar.BackColor = vbGreen
            lblScanQ.Caption = txtItemID
            lblCountLast.Caption = "Count (" & !ScanQ & ")"
        End If
    End With

    'show the updates in sub form
    'ScanBatchTotalsSub.Requery

End Sub

Private Sub txtCountLast_Click()
    ShowScanBatchLineCounts
End Sub


Private Sub txtScanCounts_Click()
    DoCmd.OpenForm "ScanBatch1"

End Sub

Private Sub txtScanQ_Click()
    txtCountLast_Click
End Sub

Private Sub ShowScanBatchLineCounts()

    Set mrsSBLC = Me.RecordsetClone
    With mrsSBLC
        txtCodeLast = txtCode
        .FindFirst "[ItemID] = " & txtItemIDLast
        If .EOF Then
            Beep
            Beep
            Beep
        Else
            Me.Bookmark = .Bookmark
            lstEntries.Visible = False
'            ScanBatchLineCountsSub.Visible = True
'            ScanBatchLineCountsSubLabel.Visible = True
'            ScanBatchLineCountsSub.SetFocus
        End If
    End With
End Sub

Private Sub BlankInfoTextBoxes()
    txtCode = ""
    txtDescription = ""
    txtPrice = ""
    txtCount = ""
End Sub
