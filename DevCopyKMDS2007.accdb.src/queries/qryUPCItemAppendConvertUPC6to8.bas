Operation =4
Option =0
Where ="(((CDbl([id]))<>[ItemID]) AND ((Len([ID]))=6))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Name ="UPCItem.ID"
    Expression ="upcecheckdigit([id])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
    ColumnsShown =579
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
