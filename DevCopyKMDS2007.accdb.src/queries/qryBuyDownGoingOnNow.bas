Operation =1
Option =0
Where ="(((BuyDown.BeginDate)<Now()) AND ((BuyDown.EndDate)>Now()))"
Begin InputTables
    Name ="BuyDownItem"
    Name ="BuyDown"
End
Begin OutputColumns
    Expression ="BuyDownItem.ItemsCode"
    Expression ="BuyDownItem.SpcPrice"
    Expression ="BuyDown.BeginDate"
    Expression ="BuyDown.EndDate"
End
Begin Joins
    LeftTable ="BuyDownItem"
    RightTable ="BuyDown"
    Expression ="BuyDownItem.BuyDownID = BuyDown.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="BuyDown.BeginDate"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDown.EndDate"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =207
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =210
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =210
        Top =0
        Name ="BuyDown"
        Name =""
    End
End
