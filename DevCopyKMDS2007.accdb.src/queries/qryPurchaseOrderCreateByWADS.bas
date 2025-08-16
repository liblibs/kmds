Operation =1
Option =0
Where ="(((Item.VendorID)=1))"
Begin InputTables
    Name ="Item"
    Name ="DaysSupplyGroupLevels"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.AverageDailySales"
    Expression ="DaysSupplyGroupLevels.DaysSupply"
    Alias ="OrderMax"
    Expression ="CLng(IIf(IsNull([AverageDailySales]),0,[AverageDailySales])*[DaysSupply])"
    Expression ="Item.Onhand"
    Alias ="RSOQ"
    Expression ="[OrderMax]-[Onhand]"
    Expression ="Item.VendorID"
    Alias ="SOQ"
    Expression ="(CLng([RSOQ]/[UnitsPerOrder])*[UnitsPerOrder])"
    Expression ="Item.UnitsPerOrder"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="DaysSupplyGroupLevels"
    Expression ="Item.DaysSupplyGroupID = DaysSupplyGroupLevels.GroupLevel"
    Flag =1
End
Begin OrderBy
    Expression ="Item.ID"
    Flag =0
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
        dbText "Name" ="Item.ID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbInteger "ColumnWidth" ="0"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderMax"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RSOQ"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerOrder"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SOQ"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysSupplyGroupLevels.DaysSupply"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.AverageDailySales"
        dbInteger "ColumnWidth" ="1770"
        dbInteger "ColumnOrder" ="3"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =165
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =273
        Top =7
        Right =544
        Bottom =121
        Top =0
        Name ="DaysSupplyGroupLevels"
        Name =""
    End
End
