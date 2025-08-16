Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub CatchAllItemID_DblClick(Cancel As Integer)
    Dim RSItem As Recordset
    Dim rs As Object
    Dim dblID As Double
    If ID = 1 Then
        Exit Sub
    End If
    If IsNull(ID) Then
        MsgBox "ID is null"
        Exit Sub
    End If
    If MsgBox("Do you want to create a catch all item for sales department #" & ID & " " & Description & "?", vbYesNo) = vbYes Then
        Set RSItem = CurrentDb().OpenRecordset("Item")

        If gcfHandleErrors Then On Error GoTo ErrLog
        Set rs = CurrentDb().OpenRecordset("select max(ID) as MaxID from Item Where ID < 10000000000")
        rs.MoveLast
        dblID = rs!MaxID + 1
        rs.Close
        Set rs = Nothing
        With RSItem
            .AddNew
            !ID = dblID
            !Description = "CatchAll - " & Description
            !SalesDepartmentID = ID
            !Price = 0
            .Update
            .Close
        End With
        CatchAllItemID.Requery
        CatchAllItemID = dblID
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "CatchAllItemID"
    Resume ExitSub

End Sub

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True
End Sub

Private Sub Form_Current()
 '   CatchAllItemID.Requery
    Me.Parent("cmdZeroNegativeOnhands").Visible = False
    If Me.NewRecord Then
        Me.Parent("cmdSalesdepartmentGroupID").Visible = False

    Else
        Me.Parent("cmdSalesdepartmentGroupID").Caption = "Move Department #" & Me.ID
        Me.Parent("cmdSalesdepartmentGroupID").Tag = Me.ID
        Me.Parent("cmdSalesdepartmentGroupID").Visible = True
        If ID > 2 Then
            Me.Parent("cmdZeroNegativeOnhands").Caption = "Reset negative onhands for Dpt #" & ID
            Me.Parent("cmdZeroNegativeOnhands").Visible = True
            Me.Parent("cmdZeroNegativeOnhands").Tag = ID

        End If
    End If

End Sub

Private Sub Form_Load()
    Me.OrderBy = "ID"
    Me.OrderByOn = True
    gbRunMailChanges = False
    ID.ColumnWidth = -2
    Description.ColumnWidth = -2
    CatchAllItemID.ColumnWidth = -2

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If gbRunMailChanges Then MailChanges False

End Sub
