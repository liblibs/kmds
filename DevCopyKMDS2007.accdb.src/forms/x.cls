Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim strTemp As String


Private Sub CurrMM_Click()
    strTemp = ShowDetail(Format(Now, "mm"), txtCurYear.BackColor)
End Sub

Private Sub Form_Load()
Dim nI As Integer

If gcfHandleErrors Then On Error GoTo ErrLog
For nI = 1 To 12
    Me("M" & nI).BackColor = txtCurYear.BackColor
Next nI

For nI = CInt((Format(Now, "MM")) - 1) To 11
    Me("M" & nI + 1).BackColor = txtPrevYear.BackColor
Next nI

lblCurrMM.Caption = Format(Now, "mmm-yyyy")
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub M1_Click()
    strTemp = ShowDetail("01", M1.BackColor)
End Sub

Private Sub M10_Click()
    strTemp = ShowDetail("10", M10.BackColor)
End Sub

Private Sub M11_Click()
    strTemp = ShowDetail("11", M11.BackColor)
End Sub

Private Sub M12_Click()
    strTemp = ShowDetail("12", M12.BackColor)
End Sub

Private Sub M2_Click()
    strTemp = ShowDetail("02", M2.BackColor)
End Sub

Private Sub M3_Click()
    strTemp = ShowDetail("03", M3.BackColor)
End Sub

Private Sub M4_Click()
    strTemp = ShowDetail("04", M4.BackColor)
End Sub

Private Sub M5_Click()
    strTemp = ShowDetail("05", M5.BackColor)
End Sub

Private Sub M6_Click()
    strTemp = ShowDetail("06", M6.BackColor)
End Sub

Private Sub M7_Click()
    strTemp = ShowDetail("07", M7.BackColor)
End Sub

Private Sub M8_Click()
    strTemp = ShowDetail("08", M8.BackColor)
End Sub

Private Sub M9_Click()
    strTemp = ShowDetail("09", M9.BackColor)
End Sub


Public Function ShowDetail(mm As String, BackColor As Long) As String
Dim datBegin As Date
Dim datEnd As Date
Dim stDocName As String

    ShowDetail = mm & "/01/"

    Select Case BackColor
    Case txtCurYear.BackColor
        ShowDetail = ShowDetail & Format(Now, "yyyy")
    Case Else
        ShowDetail = ShowDetail & Format(DateAdd("yyyy", -1, Now), "yyyy")
    End Select

    datBegin = CDate(ShowDetail)
    datEnd = DateAdd("d", -1, DateAdd("m", 1, datBegin)) & " 23:59:59"
    'datBegin = DateAdd("d", -1, datBegin) & " 23:59:59"
    txtDateBegin = datBegin
    txtDateEnd = datEnd
    'If not gcfHandleErrors then Debug.Print Code & " " & TransType & " " & txtDateBegin & " " & txtDateEnd
    If Not gcfHandleErrors Then Debug.Print Me.Parent.Name
    Select Case Me.Parent.Name
    Case "Item"
        DoCmd.OpenForm "ItemTransactionsByMonth", acNormal ' , acFormDS  ', , , acFormReadOnly, acDialog

        'DoCmd.OpenQuery "qryInvoLineByItemByTransByMonth", acViewNormal, acReadOnly
    Case "PurchaseOrdersOnFile"
        DoCmd.OpenForm "ItemTransactionsByMonthPO", acNormal  ', , , acFormReadOnly, acDialog
        'DoCmd.OpenQuery "qryInvoLineByItemByTransByMonthPO", acViewNormal, acReadOnly
    Case Else
    End Select

End Function
