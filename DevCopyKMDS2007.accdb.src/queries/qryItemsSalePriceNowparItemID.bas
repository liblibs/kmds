Operation =1
Option =0
Where ="(((qryItemsSalePriceNow.ItemID)=[parItemID]))"
Begin InputTables
    Name ="qryItemsSalePriceNow"
End
Begin OutputColumns
    Expression ="qryItemsSalePriceNow.ItemID"
    Expression ="qryItemsSalePriceNow.StartDate"
    Expression ="qryItemsSalePriceNow.EndDate"
    Expression ="qryItemsSalePriceNow.SalePrice"
End
Begin Parameters
    Name ="parItemID"
    Flag =16
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
        dbText "Name" ="qryItemsSalePriceNow.ItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-241
    Top =35
    Right =984
    Bottom =712
    Left =-1
    Top =-1
    Right =1193
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =356
        Bottom =122
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
End
