Operation =3
Name ="UPCItem"
Option =0
Where ="(((upcatoe([id]))<>[id]))"
Begin InputTables
    Name ="qryUPCItemConvertUPCAtoUPCEqry"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="qryUPCItemConvertUPCAtoUPCEqry.ItemID"
    Alias ="NewID"
    Name ="ID"
    Expression ="upcatoe([id])"
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
dbBoolean "UseTransaction" ="-1"
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
        dbText "Name" ="NewID"
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
