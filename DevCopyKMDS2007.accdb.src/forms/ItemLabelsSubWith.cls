Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.Filter = "((Not ItemLabels.Labels=0))"
    Me.FilterOn = True
    Me.DatasheetBackColor = vbYellow
    Labels.ColumnWidth = 500
    ID.ColumnWidth = 1500
End Sub
