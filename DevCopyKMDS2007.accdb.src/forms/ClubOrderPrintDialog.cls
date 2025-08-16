Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdEmailCSV_Click()

Dim fso As FileSystemObject
Dim f As Folder
Dim ts As Object
Dim fldr As String
Dim strOutput As String
Dim coleMail As New Collection


Dim stDocName As String

If gcfHandleErrors Then On Error GoTo ErrLog


    strOutput = ExportClubInvoiceToCSV()
    If strOutput <> "NoOutput" Then
        coleMail.Add strOutput
        KMDSeMail EMail, coleMail, "Invoice #" & txtInvoiceNo & " in .csv format from " & gstrStoreName
        MsgBox "Complete"
    Else
        MsgBox "ExportClubInvoiceToCSV returned NoOutput"
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Mail Invoice as .csv (excel)"

    Resume ExitSub

End Sub

Private Sub cmdFaxOG_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdFax_Click

    Dim stDocName As String
    If Me.Dirty Then
        Me.Dirty = False
    End If

    Select Case ClubOrderGuideReport
    Case "ClubOrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesByCategory"
        Else
            stDocName = "ClubOrderGuides"
        End If

    Case "SingleColumn OrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesWithShotPrice"
        Else
            stDocName = "ClubOrderGuidesWithShotPrice"
        End If
    Case Else
        MsgBox "Order Guide report not defined for " & ClubOrderGuideReport
    End Select


    ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo

    FaxCopy gstrLastArchivedReport, Phone2, "Club Order Guide ", txtClubName, "KMDS FaxCOPY"
    gstrLastArchivedReport = ""
Exit_cmdFax_Click:
    Exit Sub

Err_cmdFax_Click:
    MsgBox Err.Description
    Resume Exit_cmdFax_Click
End Sub

Private Sub cmdPrintPreviewInvoice_Click()
    PrintPreView ("ClubInvoice")

End Sub

Private Sub cmdPrintPreviewInvoiceCombined_Click()
Dim stDocName As String

If gcfHandleErrors Then On Error GoTo ErrLog

    If InvoiceByCategory Then
        stDocName = "ClubInvoiceLiquorAndNonLiquorByCategory"
    Else
        stDocName = "ClubInvoiceLiquorAndNonLiquor"
    End If
    PrintPreView (stDocName)
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub

Private Sub cmdPrintPreviewInvoiceStoreCopy_Click()
    PrintPreView ("ClubInvoiceStoreCopy")

End Sub

Private Sub cmdPrintPreviewOrderGuide_Click()
Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case ClubOrderGuideReport
    Case "ClubOrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesByCategory"
        Else
            stDocName = "ClubOrderGuides"
        End If
        PrintPreView (stDocName)

    Case "SingleColumn OrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesWithShotPrice"
        Else
            stDocName = "ClubOrderGuidesWithShotPrice"
        End If
        PrintPreView (stDocName)
    Case "ClubOrderGuidesWithSales"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesWithSales"
        Else
            stDocName = "ClubOrderGuidesWithSales"
        End If
        PrintPreView (stDocName)

    Case Else
        MsgBox "Order Guide report not defined for " & ClubOrderGuideReport
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Description
    Resume ExitSub


End Sub

Private Sub cmdPrintPreviewPickList_Click()

    Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog
    If PickListByCategory Then
        stDocName = "ClubPickListByCategory"
    Else
        stDocName = "ClubPickList"
    End If
    PrintPreView (stDocName)
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub

Private Sub Form_Close()
    DoCmd.Close acForm, "ClubOrder"

End Sub

Private Sub Form_Current()
Dim stDocName As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    cmdMailInvoice.Visible = False
    cmdEmailCSV.Visible = False
    lblEmailCsv.Visible = False
    lblEmailPdf.Visible = False


    cmdMailOrderGuide.Visible = False
    If Not IsNull(EMail) And EMail <> "" Then
        cmdMailInvoice.Visible = True
        cmdMailOrderGuide.Visible = True
        cmdEmailCSV.Visible = True
        lblEmailCsv.Visible = True
        lblEmailPdf.Visible = True
    End If

    cmdFaxOG.Visible = False
    cmdFax2.Visible = False
    If Not IsNull(Phone2) And Phone2 <> "" Then
        cmdFaxOG.Visible = True
        cmdFax2.Visible = True
    End If

    If (txtInvoiceNo = "PrintOrderGuide") Then
         txtOrderGuides = 1
         txtInvoices = 0
         txtPickLists = 1
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesByCategory"
        Else
            stDocName = "ClubOrderGuides"
        End If
        DoCmd.OpenReport stDocName, acViewPreview
        DoCmd.PrintOut , , , , txtOrderGuides
        DoCmd.Close acReport, "ClubOrderGuides"
        If gcfHandleErrors Then On Error Resume Next
        cmdClose_Click
        Exit Sub
    End If




    'default to the number of documents in each club's record
    txtOrderGuides = Nz(OrderGuides, 0)
    txtInvoices = Nz(Invoices, 0)
    txtStoreInvoices = Nz(StoreInvoices, 0)
    'added 2007-06-23

    If txtInvoices > 1 Then txtInvoices = 1
    txtInvoicesCombined = Nz(InvoicesCombined, 0)
    txtPickLists = Nz(Picklists, 0)
    txtPOSPickLists = Nz(PosPicks, 0)
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "ClubOrderPrintDialog:Form_Current"
    Resume ExitSub
End Sub

Private Sub cmdPrint_Click()
    If gcfHandleErrors Then On Error GoTo Err_cmdPrint_Click

    Dim stDocName As String

    If Me.Form.Dirty Then Me.Form.Dirty = False
    'assign the invoice number here
'    If IsNull(txtInvoiceNo) Or txtInvoiceNo = "" Then
'        AssignClubInvoiceNO CLng(txtOrdheadID)
'
'    End If

    'pick list print
    If (txtInvoiceNo <> "PrintOrderGuide") Then

        If txtPickLists > 0 Then
            If txtPickLists > 4 Then txtPickLists = 4
            If PickListByCategory Then
                stDocName = "ClubPickListByCategory"
            Else
                stDocName = "ClubPickList"
            End If
            ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo
            DoCmd.OpenReport stDocName, acViewPreview
            DoCmd.PrintOut , , , , txtPickLists
            DoCmd.Close acReport, stDocName
        End If

        'invoice print

        stDocName = "ClubInvoice"
        ArchiveReportDocument stDocName, stDocName, txtClubName, txtInvoiceNo
        ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo
        If txtInvoices > 0 Then
            If txtInvoices > 5 Then txtInvoices = 5
            DoCmd.OpenReport stDocName, acViewPreview
            DoCmd.PrintOut , , , , txtInvoices
            DoCmd.Close acReport, stDocName
        End If

        'added 2007-12-28
        'added 2007-06-23
        stDocName = "ClubInvoiceStoreCopy"
        ArchiveReportDocument stDocName, stDocName, txtClubName, txtInvoiceNo
        ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo
        If txtStoreInvoices > 0 Then
            If txtStoreInvoices > 5 Then txtStoreInvoices = 5
            DoCmd.OpenReport stDocName, acViewPreview
            DoCmd.PrintOut , , , , txtStoreInvoices
            DoCmd.Close acReport, stDocName
        End If


NoClubInvoice:
        'invoice print
        stDocName = "ClubInvoiceLiquorAndNonLiquor"
        If InvoiceByCategory Then
            stDocName = "ClubInvoiceLiquorAndNonLiquorByCategory"
        Else
            stDocName = "ClubInvoiceLiquorAndNonLiquor"
        End If
        ArchiveReportDocument stDocName, stDocName, txtClubName, txtInvoiceNo
        ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo
        If txtInvoicesCombined > 0 Then
            If txtInvoicesCombined > 5 Then txtInvoicesCombined = 5
            DoCmd.OpenReport stDocName, acViewPreview
            DoCmd.PrintOut , , , , txtInvoicesCombined
            DoCmd.Close acReport, stDocName
        End If


        'Order Guide print
        If txtOrderGuides > 0 Then
            If txtOrderGuides > 4 Then txtOrderGuides = 4

            Select Case ClubOrderGuideReport
            Case "ClubOrderGuides"
                If OrderGuideByCategory Then
                    stDocName = "ClubOrderGuidesByCategory"
                Else
                    stDocName = "ClubOrderGuides"
                End If

            Case "SingleColumn OrderGuides"
                If OrderGuideByCategory Then
                    stDocName = "ClubOrderGuidesWithShotPrice"
                Else
                    stDocName = "ClubOrderGuidesWithShotPrice"
                End If
            Case "ClubOrderGuidesWithSales"
                If OrderGuideByCategory Then
                    stDocName = "ClubOrderGuidesWithSales"
                Else
                    stDocName = "ClubOrderGuidesWithSales"
                End If
            Case Else
                MsgBox "Order Guide report not defined for " & ClubOrderGuideReport
                GoTo NoOrderGuide
            End Select



            ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo
            DoCmd.OpenReport stDocName, acViewPreview
            DoCmd.PrintOut , , , , txtOrderGuides
            DoCmd.Close acReport, stDocName
        End If
NoOrderGuide:

        If txtPOSPickLists > 0 Then
            PrintPOSPickList
        End If
    End If

    DoCmd.Close acForm, Me.Name

Exit_cmdPrint_Click:

    Exit Sub

Err_cmdPrint_Click:
    Select Case Err.Number
    Case 2501 'no data to report
        If stDocName = "ClubInvoice" Then 'jump over the club invoices
            Resume NoClubInvoice
        Else
            Resume Next
        End If
    Case Else
        EMsgBox Me, "cmdPrint" & vbCrLf & Err.Description & " " & Err.Description
    End Select
    Resume Exit_cmdPrint_Click

End Sub

Private Sub PrintPreView(stDocName As String)

If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click
    'assign the invoice number here
    'If IsNull(txtInvoiceNo) Or txtInvoiceNo = "" Then
    '    AssignClubInvoiceNO CLng(txtOrdheadID)
        'Forms("ClubOrderPrintDialog").Requery
        'Me.Refresh
    'End If

    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPrintPreview_Click:


    Exit Sub

Err_cmdPrintPreview_Click:
    Select Case Err.Number
    Case 2501 'no data to report
    Case Else
        EMsgBox Me.Form, "PrintPreview" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume Exit_cmdPrintPreview_Click

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

Private Sub Form_Load()
SetFormSize Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strPickListByCategory As String

    txtInvoiceNo = Me.OpenArgs
    If IsNull(txtInvoiceNo) Then
        MsgBox "Unknown Invoice Number, try printing again!"
        Cancel = True
    End If
    strPickListByCategory = GetKMDSSettings("PickListByCategory", "False")
    If Trim(strPickListByCategory) = "False" Then
        PickListByCategory = False
    Else
        PickListByCategory = True
    End If


End Sub

Private Sub PrintPOSPickList()

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim strMsg As String

Dim nBottleCount As Integer

Dim n750 As Integer
Dim n175 As Integer
Dim nOther As Integer

If gcfHandleErrors Then On Error GoTo ErrLog
     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubInvoiceAndSysfilePOSPrinter"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parInvoiceNo", adInteger, adParamInput)
        .Parameters("parInvoiceNo") = txtInvoiceNo
        Set rstemp = .Execute
        With rstemp
            If .RecordCount Then
                strMsg = strMsg & "POS Pick List" & vbCrLf & vbCrLf
                strMsg = strMsg & !ClubName & vbCrLf & vbCrLf
                strMsg = strMsg & "Order #" & !InvoiceNo & vbCrLf
                Do While Not .EOF
                    nBottleCount = nBottleCount + !QuantityShipped
                    Select Case !SizeDesc
                    Case "750ML"
                        n750 = n750 + !QuantityShipped
                    Case "1.75L"
                        n175 = n175 + !QuantityShipped
                    Case Else
                        nOther = nOther + !QuantityShipped
                    End Select
                    strMsg = strMsg & Format(CStr(!QuantityShipped), "@@@@") & " ______ " & !ItemID & "   " & Left(!Description, 20) & vbCrLf & vbCrLf
                    .MoveNext
                Loop
            strMsg = strMsg & "*SIZE             Cases      Break" & vbCrLf
            strMsg = strMsg & "__________________________________" & vbCrLf
            strMsg = strMsg & "750ML:  " & Format(CStr(n750), "@@@@")
            strMsg = strMsg & "  " & Format(CStr((n750 \ 12)), "@@@@@@@@")
            strMsg = strMsg & "  " & Format(CStr((n750 - ((n750 \ 12) * 12))), "@@@@@@@@")
            strMsg = strMsg & vbCrLf

            strMsg = strMsg & "1.75L:  " & Format(CStr(n175), "@@@@")
            strMsg = strMsg & "  " & Format(CStr((n175 \ 6)), "@@@@@@@@")
            strMsg = strMsg & "  " & Format(CStr((n175 - ((n175 \ 6) * 6))), "@@@@@@@@")
            strMsg = strMsg & vbCrLf

            strMsg = strMsg & "Other:  " & Format(CStr(nOther), "@@@@") & vbCrLf
            strMsg = strMsg & "__________________________________" & vbCrLf
            strMsg = strMsg & "Total:  " & Format(CStr(nBottleCount), "@@@@") & vbCrLf
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & vbCrLf
            strMsg = strMsg & vbCrLf
           End If
            .Close
        End With
        Set rstemp = Nothing
    End With
    Set comQDF = Nothing
    HardCopy strMsg, "POS Pick List"
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "PrintPOSPickList" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub Command34_Click()
If gcfHandleErrors Then On Error GoTo Err_Command34_Click

    Dim stDocName As String

    stDocName = "ClubInvoiceLiquorAndNonLiquor"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command34_Click:
    Exit Sub

Err_Command34_Click:
    MsgBox Err.Description
    Resume Exit_Command34_Click

End Sub

Private Sub cmdMailInvoice_Click()

Dim fso As FileSystemObject
Dim f As Folder
Dim ts As Object
Dim fldr As String
Dim strOutput As String
Dim coleMail As New Collection


Dim stDocName As String

If gcfHandleErrors Then On Error GoTo ErrLog

    stDocName = "ClubInvoiceLiquorAndNonLiquor"
    If InvoiceByCategory Then
        stDocName = "ClubInvoiceLiquorAndNonLiquorByCategory"
    Else
        stDocName = "ClubInvoiceLiquorAndNonLiquor"
    End If

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"

    Set fso = New FileSystemObject

    fldr = "\\" & gstrCashsysInsLoc & "\KMDS\Clubs\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    strOutput = fldr & txtClubName & " - " & txtInvoiceNo & ".pdf"

    If fso.FileExists(strOutput) Then
        fso.DeleteFile strOutput
    End If

    DoCmd.OutputTo acOutputReport, stDocName, acFormatPDF, strOutput
    coleMail.Add strOutput

    KMDSeMail EMail, coleMail, "Invoice #" & txtInvoiceNo & " in .pdf format from " & gstrStoreName
    MsgBox "Complete"

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Mail Invoice"

    Resume ExitSub
End Sub

Private Sub cmdFax2_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdFax_Click

    Dim stDocName As String
    If Me.Dirty Then
        Me.Dirty = False
    End If

    stDocName = "ClubInvoiceLiquorAndNonLiquor"
    If InvoiceByCategory Then
        stDocName = "ClubInvoiceLiquorAndNonLiquorByCategory"
    Else
        stDocName = "ClubInvoiceLiquorAndNonLiquor"
    End If

    ArchiveReport stDocName, txtClubName & " " & txtInvoiceNo

    FaxCopy gstrLastArchivedReport, Phone2, "Club Invoice " & txtInvoiceNo, txtClubName, "KMDS FaxCOPY"
    gstrLastArchivedReport = ""
Exit_cmdFax_Click:
    Exit Sub

Err_cmdFax_Click:
    MsgBox Err.Description
    Resume Exit_cmdFax_Click

End Sub
Private Sub cmdMailOrderGuide_Click()

Dim stDocName As String
Dim fso As FileSystemObject
Dim f As Folder
Dim ts As Object
Dim fldr As String
Dim strOutput As String
Dim coleMail As New Collection


    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case ClubOrderGuideReport
    Case "ClubOrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesByCategory"
        Else
            stDocName = "ClubOrderGuides"
        End If
        'PrintPreView (stDocName)

    Case "SingleColumn OrderGuides"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesWithShotPrice"
        Else
            stDocName = "ClubOrderGuidesWithShotPrice"
        End If
        'PrintPreView (stDocName)
    Case "ClubOrderGuidesWithSales"
        If OrderGuideByCategory Then
            stDocName = "ClubOrderGuidesWithSales"
        Else
            stDocName = "ClubOrderGuidesWithSales"
        End If
        'PrintPreView (stDocName)

    Case Else
        MsgBox "Order Guide report not defined for " & ClubOrderGuideReport
    End Select

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"

    Set fso = New FileSystemObject

    fldr = "\\" & gstrCashsysInsLoc & "\KMDS\Clubs\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    strOutput = fldr & txtClubName & ".pdf"

    If fso.FileExists(strOutput) Then
        fso.DeleteFile strOutput
    End If


    DoCmd.OutputTo acOutputReport, stDocName, acFormatPDF, strOutput
    coleMail.Add strOutput

    KMDSeMail EMail, coleMail, "Order guide from " & gstrStoreName
    MsgBox "Complete"

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Mail Order guide"

    Resume ExitSub

End Sub



Private Sub PickListByCategory_Click()
    If PickListByCategory Then
        SetKMDSSettings "PickListByCategory", "True"
    Else
        SetKMDSSettings "PickListByCategory", "False"
    End If

End Sub
