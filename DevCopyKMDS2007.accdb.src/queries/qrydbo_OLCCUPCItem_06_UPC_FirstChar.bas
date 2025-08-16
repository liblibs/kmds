Operation =1
Option =0
Where ="(((Len([UPCCode]))=14))"
Begin InputTables
    Name ="OLCCScanCode"
End
Begin OutputColumns
    Expression ="OLCCScanCode.UPCCode"
    Alias ="lid"
    Expression ="Len([UPCCode])"
    Alias ="UPC10"
    Expression ="Mid([upccode],4,10)"
    Alias ="FirstChar"
    Expression ="Mid([upccode],3,1)"
    Alias ="upc12"
    Expression ="checkdigit11([FirstChar] & [UPC10])"
    Expression ="OLCCScanCode.Code"
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
Begin
    Begin
        dbText "Name" ="OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstChar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upc12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPC10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =62
    Top =98
    Right =1485
    Bottom =795
    Left =-1
    Top =-1
    Right =1400
    Bottom =411
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="OLCCScanCode"
        Name =""
    End
End
