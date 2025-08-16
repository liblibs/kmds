Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cboID_AfterUpdate()
   ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Me![cboID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub

End Sub



Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec
    Desc.SetFocus
Exit_cmdfrmAddNew:
    Exit Sub

Err_cmdfrmAddNew:
    MsgBox Err.Description
    Resume Exit_cmdfrmAddNew
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70



ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501
        MsgBox "There are related items for this record"
    Case Else
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    End Select
    Resume ExitSub
End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmAddNew.Visible = True
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

    'refresh combo box,
'   cboID.Requery
'   cboID = ID
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmSaveRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmUndo_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    cmdfrmAddNew.Visible = True
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub

Private Sub cmdUndoPostedPO_Click()
Dim strMsg As String
Dim nOrderID As Long
Dim stDocName As String

If gcfHandleErrors Then On Error GoTo ErrLog

    strMsg = strMsg & "UNDO Posted Purchase Order with InvoHeadID = " & txtInvoiceHeadID
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "This will create 2 purchase orders; one is negative, the other is positive." & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "You will need to post the negative one in Menu option C-1 to 'undo' this order." & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "The positive one can be edited and posted when the order is received." & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "" & vbCrLf
    strMsg = strMsg & "Do you want to undo this Purchase Order?"
    If MsgBox(strMsg, vbYesNo) = vbYes Then
        'step 1 create blank purchase order for negative order
        nOrderID = CreatePurchaseOrderBlankOrder(ID)
        txtNewPO = nOrderID

        'step 2 append blank purchase order for negative order
        SetOptions False
        stDocName = "qryUndoPostedPONegative"
        DoCmd.OpenQuery stDocName, acNormal, acEdit

        'step 3 create blank purchase order for negative order
        nOrderID = CreatePurchaseOrderBlankOrder(ID)
        txtNewPO = nOrderID

        'step 4 append blank purchase order for negative order
        stDocName = "qryUndoPostedPOPositive"
        DoCmd.OpenQuery stDocName, acNormal, acEdit
        SetOptions True
        strMsg = "The orders have been created!" & vbCrLf & vbCrLf
        strMsg = strMsg & "You will need to post the negative one to 'undo' this order." & vbCrLf
        strMsg = strMsg & "" & vbCrLf
        MsgBox strMsg
        DoCmd.Close
        DoCmd.OpenForm "PurchaseOrderCreate", acNormal
        Exit Sub
    End If

ExitSub:
    SetOptions True
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdUndoPostedPO_Click"
    Resume ExitSub

End Sub

Private Sub Command32_Click()
If gcfHandleErrors Then On Error GoTo Err_Command32_Click
    If Form.Dirty Then
        Form.Dirty = False
    End If
    lblPleaseWait.Visible = True
    Me.Repaint
    SetOptions False

    AssignDaysSupplyGroupIDs
    lblPleaseWait.Visible = False
Exit_Command32_Click:
    SetOptions True

    DoCmd.Close

    Exit Sub

Err_Command32_Click:
    MsgBox Err.Description
    Resume Exit_Command32_Click
End Sub

Private Sub Desc_BeforeUpdate(Cancel As Integer)
    'since club/vendor names are used in the file structure
    'they can not have illegal file names!Dim ni As Integer
    Dim nI As Integer

    For nI = 1 To Len(Desc)
        If InStr(gstrIllegalFname, Mid(Desc, nI, 1)) Then
            MsgBox "this name can not contain any of these characters " & vbCrLf & gstrIllegalFname
            Cancel = True
            Exit Sub
        End If
    Next nI

End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    If Not IsNull(cmbItemsSupplied.Column(1)) Then
        Cancel = True
    End If

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.NewRecord Then
        Desc.SetFocus
    End If
    ID.Visible = Me.NewRecord
    cboID.Visible = Not Me.NewRecord
    cboID = ID


    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord
    If Not Me.NewRecord And optBasePO = -1 Then ' obsolete
        Desc.BackColor = vbRed
        Contact.BackColor = vbRed
        Label73.ForeColor = vbRed
    Else
        Desc.BackColor = Address1.BackColor
        Contact.BackColor = Address1.BackColor
        Label73.ForeColor = Label50.ForeColor
    End If


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub



End Sub

Private Sub Form_Delete(Cancel As Integer)
    If Not IsNull(cmbItemsSupplied.Column(1)) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim nCode As Long
Dim rs As Object

    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3146 'description = blank
        Me.Undo
        DoCmd.GoToRecord , , acPrevious
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Sizes:Form_Error"
    Resume ExitSub

End Sub




Private Sub Form_Load()
    cboID.ColumnHidden = True
    ID.ColumnHidden = True
    SetupDate.ColumnHidden = True
    ChangeDate.ColumnHidden = True
    Text114.ColumnHidden = True
    Text118.ColumnHidden = True
    Text120.ColumnHidden = True
    txtInvoiceHeadID.ColumnHidden = True
    txtNewPO.ColumnHidden = True

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If Me.OpenArgs <> "" Then
            Me.Desc.SetFocus
            Me.Desc.Text = Me.OpenArgs
        End If
    End If
End Sub



Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub cmdVendorPOHistory_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdVendorPOHistory_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "VendorPOHistory"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdVendorPOHistory_Click:
    Exit Sub

Err_cmdVendorPOHistory_Click:
    MsgBox Err.Description
    Resume Exit_cmdVendorPOHistory_Click

End Sub
Private Sub cmdAverageDaySupply_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAverageDaySupply_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    If Form.Dirty Then
        Form.Dirty = False
    End If
    stDocName = "AverageDaySupply"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    cmdRefresh.Visible = True
    cmdRefresh.SetFocus
    'cmdAverageDaySupply.Visible = False
Exit_cmdAverageDaySupply_Click:
    Exit Sub

Err_cmdAverageDaySupply_Click:
    MsgBox Err.Description
    Resume Exit_cmdAverageDaySupply_Click

End Sub
Private Sub cmdRefresh_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdRefresh_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    'cmdAverageDaySupply.Visible = True
    cmdfrmClose.SetFocus
    'cmdAverageDaySupply.SetFocus
    cmdRefresh.Visible = False

Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click

End Sub
Private Sub cmdUpdateADS_Click()

    AssignDaysSupplyGroupIDs
Exit_cmdUpdateADS_Click:
    Exit Sub

Err_cmdUpdateADS_Click:
    MsgBox Err.Description
    Resume Exit_cmdUpdateADS_Click

End Sub
Private Sub cmdFixVendorNo_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdFixVendorNo_Click

    Dim stDocName As String
    SetOptions False
    stDocName = "qryInvoLineFixPOVendorLines"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    SetOptions True

Exit_cmdFixVendorNo_Click:
    Exit Sub

Err_cmdFixVendorNo_Click:
    MsgBox Err.Description
    Resume Exit_cmdFixVendorNo_Click

End Sub
Private Sub cmdUndoNeg_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdUndoNeg_Click

    Dim stDocName As String

    stDocName = "qryUndoPostedPO1"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_cmdUndoNeg_Click:
    Exit Sub

Err_cmdUndoNeg_Click:
    MsgBox Err.Description
    Resume Exit_cmdUndoNeg_Click

End Sub
Private Sub cmdPrintPreview_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click

    Dim stDocName As String

    stDocName = "PurchaseOrderWithPriceAndCostVendorScr"
    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPrintPreview_Click:
    Exit Sub

Err_cmdPrintPreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintPreview_Click

End Sub
'------------------------------------------------------------
' cmdfrmDeleteRecord_Click
'
'------------------------------------------------------------
Private Sub xcmdfrmDeleteRecord_Click()
On Error GoTo cmdfrmDeleteRecord_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.NewRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.NewRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.NewRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If
    Call SetRunMailChangesTrue


cmdfrmDeleteRecord_Click_Exit:
    Exit Sub

cmdfrmDeleteRecord_Click_Err:
    MsgBox Error$
    Resume cmdfrmDeleteRecord_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmSaveRecord_Click
'
'------------------------------------------------------------
Private Sub xcmdfrmSaveRecord_Click()
On Error GoTo cmdfrmSaveRecord_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If


cmdfrmSaveRecord_Click_Exit:
    Exit Sub

cmdfrmSaveRecord_Click_Err:
    MsgBox Error$
    Resume cmdfrmSaveRecord_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmUndo_Click
'
'------------------------------------------------------------
Private Sub xcmdfrmUndo_Click()
On Error GoTo cmdfrmUndo_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdUndo
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If


cmdfrmUndo_Click_Exit:
    Exit Sub

cmdfrmUndo_Click_Err:
    MsgBox Error$
    Resume cmdfrmUndo_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmClose_Click
'
'------------------------------------------------------------
Private Sub xcmdfrmClose_Click()
On Error GoTo cmdfrmClose_Click_Err

    DoCmd.Close , ""


cmdfrmClose_Click_Exit:
    Exit Sub

cmdfrmClose_Click_Err:
    MsgBox Error$
    Resume cmdfrmClose_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmAddNew_Click
'
'------------------------------------------------------------
Private Sub xcmdfrmAddNew_Click()
On Error GoTo cmdfrmAddNew_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If


cmdfrmAddNew_Click_Exit:
    Exit Sub

cmdfrmAddNew_Click_Err:
    MsgBox Error$
    Resume cmdfrmAddNew_Click_Exit

End Sub


'------------------------------------------------------------
' Command68_Click
'
'------------------------------------------------------------
Private Sub Command68_Click()
On Error GoTo Command68_Click_Err

    Call PrintPreView(Form)


Command68_Click_Exit:
    Exit Sub

Command68_Click_Err:
    MsgBox Error$
    Resume Command68_Click_Exit

End Sub


'------------------------------------------------------------
' cmdDataSheet_Click
'
'------------------------------------------------------------
Private Sub cmdDataSheet_Click()
On Error GoTo cmdDataSheet_Click_Err

    Call DatasheetView(Form)


cmdDataSheet_Click_Exit:
    Exit Sub

cmdDataSheet_Click_Err:
    MsgBox Error$
    Resume cmdDataSheet_Click_Exit

End Sub


'------------------------------------------------------------
' cmdDelete2_Click
'
'------------------------------------------------------------
Private Sub cmdDelete2_Click()
On Error GoTo cmdDelete2_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.NewRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.NewRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.NewRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdDelete2_Click_Exit:
    Exit Sub

cmdDelete2_Click_Err:
    MsgBox Error$
    Resume cmdDelete2_Click_Exit

End Sub


'------------------------------------------------------------
' cmdPrintItemsSuppliedByThisVendor_Click
'
'------------------------------------------------------------
Private Sub cmdPrintItemsSuppliedByThisVendor_Click()
On Error GoTo cmdPrintItemsSuppliedByThisVendor_Click_Err

    DoCmd.OpenReport "Vendor Supply List", acViewPreview, "", "", acNormal


cmdPrintItemsSuppliedByThisVendor_Click_Exit:
    Exit Sub

cmdPrintItemsSuppliedByThisVendor_Click_Err:
    MsgBox Error$
    Resume cmdPrintItemsSuppliedByThisVendor_Click_Exit

End Sub


'------------------------------------------------------------
' cmdAverageDailySupplyTo_1_Click
'
'------------------------------------------------------------
Private Sub cmdAverageDailySupplyTo_1_Click()
On Error GoTo cmdAverageDailySupplyTo_1_Click_Err

    DoCmd.RunSQL "UPDATE Vendor SET Vendor.AverageDaySupplyID = -1;"



cmdAverageDailySupplyTo_1_Click_Exit:
    Exit Sub

cmdAverageDailySupplyTo_1_Click_Err:
    MsgBox Error$
    Resume cmdAverageDailySupplyTo_1_Click_Exit

End Sub
