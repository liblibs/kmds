Operation =3
Name ="upcitem"
Option =0
Begin InputTables
    Name ="qryUPCItemConvertUPCAtoUPCEqry"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="qryUPCItemConvertUPCAtoUPCEqry.ItemID"
    Alias ="newid"
    Name ="ID"
    Expression ="checkdigit11(\"0\" & [id])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryUPCItemConvertUPCAtoUPCEqry.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCItemConvertUPCAtoUPCEqry.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="newid"
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
    Bottom =361
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryUPCItemConvertUPCAtoUPCEqry"
        Name =""
    End
End
