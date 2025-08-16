Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
Dim nI As Integer
nI = -2
SetupDate.ColumnWidth = nI
Customer.ColumnWidth = nI
Invoice.ColumnWidth = nI
OrderNotes.ColumnWidth = nI
Qty.ColumnWidth = 900
EMail.ColumnWidth = nI

End Sub
