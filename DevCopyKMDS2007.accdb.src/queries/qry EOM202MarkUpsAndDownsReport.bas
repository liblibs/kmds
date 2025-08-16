Operation =1
Option =0
Where ="(((Item.Cost)<>[Price]) AND ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="Old"
    Expression ="Item.Cost"
    Alias ="New"
    Expression ="Item.Price"
    Alias ="PriceDifference"
    Expression ="[Price]-[Cost]"
    Expression ="Item.Onhand"
    Alias ="OldValue"
    Expression ="[onhand]*[Cost]"
    Alias ="NewValue"
    Expression ="([onhand]*[Price])"
    Alias ="Markup"
    Expression ="IIf(([Price]>[Cost]),([Price]-[Cost])*[onhand],0)"
    Alias ="MarkDown"
    Expression ="IIf(([Price]<[Cost]),([Price]-[Cost])*[onhand],0)"
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
        dbText "Name" ="New"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="4080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Old"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbInteger "ColumnWidth" ="645"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceDifference"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Markup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MarkDown"
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
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =219
        Top =21
        Right =315
        Bottom =135
        Top =0
        Name ="Item"
        Name =""
    End
End
