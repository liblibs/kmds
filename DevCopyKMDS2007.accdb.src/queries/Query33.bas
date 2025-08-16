Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND ((Item.SetupDate)>#7/27/2015#))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.Description"
    Expression ="Item.SizeID"
    Expression ="Item.SetupDate"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1465
    Bottom =851
    Left =-1
    Top =-1
    Right =1433
    Bottom =528
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =358
        Top =0
        Name ="Item"
        Name =""
    End
End
