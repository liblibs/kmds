Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbSize_DblClick(Cancel As Integer)
    AddCurrentRecord
End Sub



Private Sub Code_DblClick(Cancel As Integer)
    AddCurrentRecord
End Sub

Public Sub AddCurrentRecord()
Dim strSQL As String
Dim nVendorID As Integer
Dim nPosition As Long

    If gcfHandleErrors Then On Error GoTo ErrLog

    'set special order items to vendor = 0 to prevent ordering!
    nVendorID = 1
    If StockStatus = "S" Then
        strSQL = "Special Order items MUST be ordered through OLAS" & vbCrLf
        strSQL = strSQL & "Are you sure you want to add " & Description & "?"
        If MsgBox(strSQL, vbYesNo) <> vbYes Then
            Exit Sub
        End If
        'this was not working good
        'nVendorID = 0
    End If

    strSQL = "INSERT INTO Item ( ID, "
    strSQL = strSQL & "Description, "
    strSQL = strSQL & "SalesDepartmentID, "

    strSQL = strSQL & "SizeID, "
    strSQL = strSQL & "UnitsPerCase, "
    strSQL = strSQL & "UnitsPerOrder, "
    strSQL = strSQL & "VendorID, "
    strSQL = strSQL & "Cost, "
    strSQL = strSQL & "Price, "
    strSQL = strSQL & "SubCategoryID, "
    strSQL = strSQL & "StockNumber ) "


    strSQL = strSQL & "SELECT " & Code & " AS ID, "
    strSQL = strSQL & "'" & Description & "' AS Description, "
    strSQL = strSQL & "1 AS SalesDepartmentID, "
    strSQL = strSQL & cmbSize & " AS SizeID, "
    strSQL = strSQL & UnitsPerCase & " AS UnitsPerCase, "
    strSQL = strSQL & UnitsPerCase & " AS UnitsPerOrder, "
    strSQL = strSQL & nVendorID & " AS VendorID, "
    strSQL = strSQL & Price & " AS Cost, "
    strSQL = strSQL & Price & " AS Price, "
    strSQL = strSQL & Nz(SubCategoryID, 0) & " AS SubCategoryID, "
    strSQL = strSQL & "'" & StockStatus & "' AS StockNumber "

    CurrentDb().Execute (strSQL)

    strSQL = "INSERT INTO UPCItem ( ID, ItemID ) "
    strSQL = strSQL & "SELECT " & Code & " AS ID, "
    strSQL = strSQL & Code & " AS ItemID"
    If gcfHandleErrors Then On Error Resume Next
    CurrentDb().Execute (strSQL)

    If Not IsNull(UPCCodeKMDS) Then
        strSQL = "INSERT INTO UPCItem ( ID, ItemID ) "
        strSQL = strSQL & "SELECT '" & CStr(UPCCodeKMDS) & "' AS ID, "
        strSQL = strSQL & Code & " AS ItemID"
        If gcfHandleErrors Then On Error Resume Next
        CurrentDb().Execute (strSQL)
    End If


    nPosition = Me.Recordset.AbsolutePosition
    Me.Requery
    Me.Recordset.AbsolutePosition = nPosition
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbCritical, "AddCurrentRecord"
    Resume ExitSub

End Sub

Private Sub Combo14_DblClick(Cancel As Integer)
AddCurrentRecord

End Sub

Private Sub Description_DblClick(Cancel As Integer)
AddCurrentRecord
End Sub



Private Sub Price_DblClick(Cancel As Integer)
AddCurrentRecord
End Sub

Private Sub StockStatus_DblClick(Cancel As Integer)
AddCurrentRecord
End Sub


Private Sub UnitsPerCase_DblClick(Cancel As Integer)
AddCurrentRecord
End Sub
