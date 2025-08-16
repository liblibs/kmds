Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND ((OLASItemDetailForOrdering.NewItemCode) Is Nu"
    "ll))"
Begin InputTables
    Name ="Item"
    Name ="OLASItemDetailForOrdering"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.ID"
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.SizeID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="OLASItemDetailForOrdering"
    Expression ="Item.ID = OLASItemDetailForOrdering.NewItemCode"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Item.AvgTurns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.QuantityAvailable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLASItemDetailForOrdering.NewItemCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-17
    Top =5
    Right =962
    Bottom =650
    Left =-1
    Top =-1
    Right =696
    Bottom =310
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =223
        Bottom =316
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =290
        Top =8
        Right =490
        Bottom =297
        Top =0
        Name ="OLASItemDetailForOrdering"
        Name =""
    End
End
