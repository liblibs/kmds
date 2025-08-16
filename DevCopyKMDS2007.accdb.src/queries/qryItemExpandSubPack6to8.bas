Operation =4
Option =0
Where ="(((Len([SubPack]))=6))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.SubPack"
    Expression ="upcecheckdigit([Subpack])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
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
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =244
        Bottom =285
        Top =0
        Name ="Item"
        Name =""
    End
End
