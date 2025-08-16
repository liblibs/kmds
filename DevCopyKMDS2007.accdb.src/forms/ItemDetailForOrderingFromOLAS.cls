Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub

'------------------------------------------------------------
' Command0_Click
'
'------------------------------------------------------------
Private Sub Command0_Click()
Dim rs
On Error GoTo Command0_Click_Err
    Set rs = CreateObject("ADODB.Recordset")
    SetOptions False
    If gcfHandleErrors Then On Error GoTo Command0_Click_Err
    ' get the totals and update the ordhead...
    Set rs = CurrentDb().OpenRecordset("qryItemDetailForOrdering_00_VerifyDownload", dbReadOnly)
    If rs.RecordCount = 0 Then
        MsgBox "Input Data Formatting Error!", vbCritical
        MsgBox "Retry Exporting report from OLAS"
        MsgBox "K:\Mail\In\Item Detail for Ordering.xls needs to be exported as:"
        MsgBox "Microsoft Excel (97-2003) DATA ONLY"
    Else
        DoCmd.OpenQuery "qryItemDetailForOrdering_01_Delete", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_02_Append", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_03_ReportDateDelete", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_04_ReportDateAppend", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_05_ReportDateUpdate", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_06_QtyPack_Update_Reset", acViewNormal, acEdit
        DoCmd.OpenQuery "qryItemDetailForOrdering_07_QtyPackNoStock_Update", acViewNormal, acEdit
        ItemDetailForOrderingFromOLAS_subMinOrder.Requery
        ItemDetailForOrderingFromOLAS_subCaseQty.Requery
        ItemDetailForOrderingFromOLAS_subReportDate.Requery
        SetOptions True
        MsgBox "Finished"
    End If
Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    Select Case Err.Number
    Case 3078
        MsgBox "K:\Mail\In\Item Detail for Ordering.XLS" & vbCrLf & "not found!", vbCritical
        MsgBox "Retry Exporting report from OLAS"
        MsgBox "K:\Mail\In\Item Detail for Ordering.xls needs to be exported as:"
        MsgBox "Microsoft Excel (97-2003) DATA ONLY"
    Case Else
        MsgBox Err.Number & " " & Error$
    End Select
    Resume Command0_Click_Exit

End Sub


'------------------------------------------------------------
' Command7_Click
'
'------------------------------------------------------------
Private Sub Command7_Click()
On Error GoTo Command7_Click_Err

    Me.Requery



Command7_Click_Exit:
    Exit Sub

Command7_Click_Err:
    MsgBox Error$
    Resume Command7_Click_Exit

End Sub


Private Sub Form_Load()
   Me.ItemDetailForOrderingFromOLAS_subMinOrder.Form![ID].ColumnWidth = 1500
   Me.ItemDetailForOrderingFromOLAS_subMinOrder.Form![UnitsPerOrder].ColumnWidth = 1500
   Me.ItemDetailForOrderingFromOLAS_subMinOrder.Form![MinOrderQty].ColumnWidth = 1500

   Me.ItemDetailForOrderingFromOLAS_subCaseQty.Form![ID].ColumnWidth = 1500
   Me.ItemDetailForOrderingFromOLAS_subCaseQty.Form![MyCaseQty].ColumnWidth = 1500
   Me.ItemDetailForOrderingFromOLAS_subCaseQty.Form![CaseQtyOLAS].ColumnWidth = 1500

   Me.ItemDetailForOrderingFromOLAS_subReportDate.Form![ReportDate].ColumnWidth = 1500
   Me.ItemDetailForOrderingFromOLAS_subReportDate.Form![Updated].ColumnWidth = 2500


End Sub
'------------------------------------------------------------
' cmdUpdCaseQuantiy_Click
'
'------------------------------------------------------------
Private Sub cmdUpdCaseQuantiy_Click()
On Error GoTo cmdUpdCaseQuantiy_Click_Err

    DoCmd.OpenQuery "qryItemDetailForOrdering_CaseQtyVariance_Update", acViewNormal, acEdit
    ItemDetailForOrderingFromOLAS_subCaseQty.Requery

cmdUpdCaseQuantiy_Click_Exit:
    Exit Sub

cmdUpdCaseQuantiy_Click_Err:
    MsgBox Error$
    Resume cmdUpdCaseQuantiy_Click_Exit

End Sub


'------------------------------------------------------------
' cmdUpdateUnitsPerOrder_Click
'
'------------------------------------------------------------
Private Sub cmdUpdateUnitsPerOrder_Click()
On Error GoTo cmdUpdateUnitsPerOrder_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="cmdUpdCaseQuantiy" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/
    ' _AXL:forms"><Statements><Action Name="OpenQuery"><Argument Name="QueryName">qryItemDetailForOrdering_CaseQtyVariance</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenQuery "qryItemDetailForOrdering_UnitsPerOrder_MinOrder_Update", acViewNormal, acEdit
    ItemDetailForOrderingFromOLAS_subCaseQty.Requery


cmdUpdateUnitsPerOrder_Click_Exit:
    Exit Sub

cmdUpdateUnitsPerOrder_Click_Err:
    MsgBox Error$
    Resume cmdUpdateUnitsPerOrder_Click_Exit

End Sub
