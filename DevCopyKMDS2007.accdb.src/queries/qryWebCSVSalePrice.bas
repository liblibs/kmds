Operation =1
Option =0
Where ="(((ItemSalePrice.EndDate)>Now()))"
Begin InputTables
    Name ="ItemSalePrice"
End
Begin OutputColumns
    Expression ="ItemSalePrice.ItemID"
    Expression ="ItemSalePrice.StartDate"
    Expression ="ItemSalePrice.EndDate"
    Expression ="ItemSalePrice.OldPrice"
    Expression ="ItemSalePrice.SalePrice"
    Expression ="ItemSalePrice.SetupDate"
    Expression ="ItemSalePrice.ChangeDate"
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
        dbText "Name" ="ItemSalePrice.EndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemSalePrice.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemSalePrice.ItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1508
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =393
        Top =0
        Name ="ItemSalePrice"
        Name =""
    End
End
