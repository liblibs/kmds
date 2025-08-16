Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Dim nI As Integer
'    If gcfHandleErrors Then On Error GoTo ErrLog
    nI = -2
    ID.ColumnHidden = True
    ItemID.ColumnWidth = 1500
    Description.ColumnWidth = 3000
    ScanCode.ColumnWidth = 2000
    OnShelf.ColumnWidth = 1000
    OnShelfDate.ColumnWidth = 1000
    InputMethod.ColumnWidth = 2000
    CountQuantity.ColumnWidth = 1000

    'Overflow.ColumnWidth = ni
    'Display.ColumnWidth = ni

   'ID.ColumnOrder = 1
    ItemID.ColumnOrder = 2
    Description.ColumnOrder = 3
    ScanCode.ColumnOrder = 4
    CountQuantity.ColumnOrder = 5
    OnShelf.ColumnOrder = 6
    OnShelfDate.ColumnOrder = 7
    InputMethod.ColumnOrder = 8

End Sub
