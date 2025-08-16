Operation =4
Option =0
Where ="(((Item.SubCategoryID)<>0) AND ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.SubCategoryID"
    Expression ="0"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
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
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =266
        Bottom =230
        Top =0
        Name ="Item"
        Name =""
    End
End
