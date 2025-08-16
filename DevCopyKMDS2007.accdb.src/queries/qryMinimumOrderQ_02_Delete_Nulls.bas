Operation =5
Option =0
Where ="(((Case_quantities.CreateDate) Is Null))"
Begin InputTables
    Name ="Case_quantities"
End
Begin OutputColumns
    Expression ="Case_quantities.*"
    Expression ="Case_quantities.CreateDate"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Case_quantities.CreateDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =120
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1466
    Bottom =483
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Case_quantities"
        Name =""
    End
End
