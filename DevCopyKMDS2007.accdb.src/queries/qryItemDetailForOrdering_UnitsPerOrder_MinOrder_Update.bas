Operation =4
Option =0
Where ="(((Item.UnitsPerOrder)<>OLASItemDetailForOrdering.MinOrderQty) And ((OLASItemDet"
    "ailForOrdering.MinOrderQty)>0))"
Begin InputTables
    Name ="Item"
    Name ="OLASItemDetailForOrdering"
End
Begin OutputColumns
    Name ="Item.UnitsPerOrder"
    Expression ="[OLASItemDetailForOrdering].[MinOrderQty]"
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
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerOrder"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.MinOrderQty"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =110
    Top =54
    Right =1108
    Bottom =716
    Left =-1
    Top =-1
    Right =966
    Bottom =310
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
        Left =492
        Top =18
        Right =636
        Bottom =300
        Top =0
        Name ="OLASItemDetailForOrdering"
        Name =""
    End
End
