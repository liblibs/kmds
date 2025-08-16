Operation =4
Option =0
Where ="(((Item.UnitsPerCase)<>[OLASItemDetailForOrdering].[BottlesPerCase]))"
Begin InputTables
    Name ="Item"
    Name ="OLASItemDetailForOrdering"
End
Begin OutputColumns
    Alias ="MyCaseQty"
    Name ="Item.UnitsPerCase"
    Expression ="[OLASItemDetailForOrdering].[BottlesPerCase]"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="OLASItemDetailForOrdering"
    Expression ="Item.ID = OLASItemDetailForOrdering.NewItemCode"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MyCaseQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CaseQtyOLAS"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =99
    Top =211
    Right =1097
    Bottom =873
    Left =-1
    Top =-1
    Right =966
    Bottom =361
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =294
        Top =0
        Name ="OLASItemDetailForOrdering"
        Name =""
    End
End
