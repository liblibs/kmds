Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.QtyPack"
    Expression ="0"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =117
    Top =184
    Right =845
    Bottom =829
    Left =-1
    Top =-1
    Right =696
    Bottom =361
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =288
        Bottom =296
        Top =0
        Name ="Item"
        Name =""
    End
End
