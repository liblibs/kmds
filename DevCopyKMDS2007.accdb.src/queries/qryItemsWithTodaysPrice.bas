Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryItemsSalePriceNow"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Price"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Alias ="DP"
    Expression ="IIf(IsNull([StartDate]),[Price],[SalePrice])"
    Expression ="qryItemsSalePriceNow.StartDate"
    Expression ="qryItemsSalePriceNow.EndDate"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
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
        dbInteger "ColumnWidth" ="600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DP"
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
        Right =134
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =195
        Top =-1
        Right =291
        Bottom =113
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
End
