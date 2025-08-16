Operation =1
Option =0
Where ="(((OrderLine.QuantityShipped)<>0) AND ((OrdHead.PaidFlag)<>0))"
Begin InputTables
    Name ="OrderLine"
    Name ="Club"
    Name ="OrdHead"
    Name ="Item"
End
Begin OutputColumns
    Expression ="OrdHead.InvoiceNo"
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="OrderLine.QuantityShipped"
    Alias ="ExtPrice"
    Expression ="([QuantityShipped]*Orderline.Price)"
    Expression ="OrdHead.DisDollars"
    Expression ="OrdHead.InvDollars"
    Alias ="NetDollars"
    Expression ="[InvDollars]-[DisDollars]"
    Expression ="OrdHead.PaidFlag"
End
Begin Joins
    LeftTable ="Club"
    RightTable ="OrdHead"
    Expression ="Club.ID = OrdHead.ClubID"
    Flag =1
    LeftTable ="OrderLine"
    RightTable ="OrdHead"
    Expression ="OrderLine.OrdheadID = OrdHead.ID"
    Flag =1
    LeftTable ="OrderLine"
    RightTable ="Item"
    Expression ="OrderLine.ItemID = Item.ID"
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
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetDollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2280"
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
    Bottom =187
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =541
        Top =1
        Right =704
        Bottom =190
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =365
        Top =74
        Right =461
        Bottom =188
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =180
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =955
        Top =21
        Right =1051
        Bottom =135
        Top =0
        Name ="Item"
        Name =""
    End
End
