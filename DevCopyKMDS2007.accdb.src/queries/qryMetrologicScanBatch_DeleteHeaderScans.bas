Operation =5
Option =0
Where ="(((MetrologicScanbatch.ScanCode)=\"-1\"))"
Begin InputTables
    Name ="MetrologicScanbatch"
End
Begin OutputColumns
    Expression ="MetrologicScanbatch.ScanCode"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =959
    Bottom =834
    Left =-1
    Top =-1
    Right =889
    Bottom =465
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MetrologicScanbatch"
        Name =""
    End
End
