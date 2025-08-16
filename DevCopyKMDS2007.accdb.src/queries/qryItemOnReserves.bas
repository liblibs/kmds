Operation =1
Option =0
Where ="(((OrderLine.QuantityShipped)<>0))"
Begin InputTables
    Name ="OrderLine"
    Name ="OrdHead"
    Name ="Club"
End
Begin OutputColumns
    Expression ="OrdHead.ID"
    Expression ="OrdHead.OrdDate"
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="OrderLine.ItemID"
    Expression ="OrderLine.QuantityShipped"
    Expression ="OrderLine.Price"
    Expression ="OrderLine.Cost"
    Expression ="OrderLine.SetupDate"
    Expression ="OrderLine.ChangeDate"
    Alias ="ExtPrice"
    Expression ="[QuantityShipped]*[Price]"
    Alias ="ExtCost"
    Expression ="[QuantityShipped]*[Cost]"
End
Begin Joins
    LeftTable ="OrderLine"
    RightTable ="OrdHead"
    Expression ="OrderLine.OrdheadID=OrdHead.ID"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="Club"
    Expression ="OrdHead.ClubID=Club.ID"
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
        dbText "Name" ="Club.ID"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.ID"
        dbInteger "ColumnWidth" ="390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.OrdDate"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.ItemID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.ChangeDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
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
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =38
        Right =156
        Bottom =152
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =356
        Top =32
        Right =452
        Bottom =146
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =551
        Top =21
        Right =647
        Bottom =135
        Top =0
        Name ="Club"
        Name =""
    End
End
