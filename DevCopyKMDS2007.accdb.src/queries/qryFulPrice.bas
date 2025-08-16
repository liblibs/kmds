Operation =1
Option =0
Where ="(((Item.Price)<>([fulprice].[Price]*0.01)))"
Begin InputTables
    Name ="Fulprice"
    Name ="Item"
    Name ="qryFulPriceMinCode"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Price"
    Alias ="WebPrice"
    Expression ="Format(([fulprice].[Price]*0.01),\"Currency\")"
    Alias ="Year"
    Expression ="Left([MInOfCode],4)"
    Alias ="Month"
    Expression ="Mid([MinOfCode],5,2)"
    Expression ="Fulprice.Price"
    Expression ="Fulprice.Code"
End
Begin Joins
    LeftTable ="Fulprice"
    RightTable ="Item"
    Expression ="Fulprice.Code = Item.ID"
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
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WebPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Month"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fulprice.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fulprice.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =240
    Top =23
    Right =1465
    Bottom =700
    Left =-1
    Top =-1
    Right =1193
    Bottom =156
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =105
        Top =0
        Name ="Fulprice"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =508
        Top =12
        Right =652
        Bottom =156
        Top =0
        Name ="qryFulPriceMinCode"
        Name =""
    End
End
