Operation =4
Option =0
Where ="(((Audit.OnShelf)<>[vwItemsWithCurrentPricing].[onshelf]))"
Begin InputTables
    Name ="Audit"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Name ="Audit.OnShelf"
    Expression ="vwItemsWithCurrentPricing.onshelf"
    Name ="Audit.VarianceAmount"
    Expression ="audit.physicalcount-vwItemsWithCurrentPricing.onshelf"
End
Begin Joins
    LeftTable ="Audit"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="Audit.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Audit.OnShelf"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbInteger "ColumnWidth" ="3255"
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
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =176
        Bottom =193
        Top =0
        Name ="Audit"
        Name =""
    End
    Begin
        Left =386
        Top =7
        Right =612
        Bottom =208
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
