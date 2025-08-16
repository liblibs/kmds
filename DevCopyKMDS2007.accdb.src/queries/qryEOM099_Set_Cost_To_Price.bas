Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Cost"
    Expression ="[ITEM].[PRICE]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =16
    Right =987
    Bottom =1302
    Left =-1
    Top =-1
    Right =932
    Bottom =438
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =347
        Top =0
        Name ="Item"
        Name =""
    End
End
