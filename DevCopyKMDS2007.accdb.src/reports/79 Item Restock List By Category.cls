Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbNoData As Boolean

Private Sub Report_Close()
    If mbNoData = True Then Exit Sub

    If MsgBox("Reset Restock Amount to 0?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        SetOptions False
        QueryExecute "qryItemsSetRestockAmountToZero"
        SetOptions True
        MsgBox "Restock Amount reset to 0.", vbInformation, "Information"
    End If
End Sub

Private Sub Report_Load()
    Me.lblTime = Now()
End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox "No data." & vbCrLf & "There are no items to restock.", vbInformation, "Information"
    Cancel = True
    mbNoData = True
End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = "ID"
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
End Sub
