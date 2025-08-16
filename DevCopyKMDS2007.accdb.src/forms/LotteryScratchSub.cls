Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
Dim nI As Integer
    nI = 1
    GameNumber.ColumnOrder = nI
    nI = nI + 1
    Description.ColumnOrder = nI
    nI = nI + 1
    Price.ColumnOrder = nI
    nI = nI + 1
    BeginMeter.ColumnOrder = nI
    nI = nI + 1
    EndMeter.ColumnOrder = nI
    nI = nI + 1
    AddMeter.ColumnOrder = nI
    nI = nI + 1
    PhysicalAmount.ColumnOrder = nI
    nI = nI + 1
    AmountSold.ColumnOrder = nI
    nI = nI + 1
    Variance.ColumnOrder = nI
    nI = nI + 1
    VarianceDollars.ColumnOrder = nI

    nI = -2
    GameNumber.ColumnWidth = nI
    Description.ColumnWidth = nI
    Price.ColumnWidth = nI
    BeginMeter.ColumnWidth = nI
    EndMeter.ColumnWidth = nI
    AddMeter.ColumnWidth = nI
    PhysicalAmount.ColumnWidth = nI
    AmountSold.ColumnWidth = nI
    Variance.ColumnWidth = nI
    VarianceDollars.ColumnWidth = nI
    SoldDollars.ColumnWidth = 0



End Sub
