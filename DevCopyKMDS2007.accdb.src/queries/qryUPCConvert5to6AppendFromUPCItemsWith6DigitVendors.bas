Operation =3
Name ="UPCConvert5to6"
Option =0
Where ="(((Len([ID]))=12) AND ((UPCItem.ID)>\"2\"))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Alias ="Vendor6"
    Name ="Vendor6Digit"
    Expression ="Left([id],6)"
    Alias ="Vendor5"
    Name ="Vendor5Digit"
    Expression ="Mid([id],2,5)"
End
Begin Groups
    Expression ="Left([id],6)"
    GroupLevel =0
    Expression ="Mid([id],2,5)"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Vendor6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor5"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
