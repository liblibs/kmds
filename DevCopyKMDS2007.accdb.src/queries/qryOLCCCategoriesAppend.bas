Operation =3
Name ="OLCC Categories"
Option =0
Begin InputTables
    Name ="OLCC Items from OLAS Website"
End
Begin OutputColumns
    Name ="Category"
    Expression ="[OLCC Items from OLAS Website].Category"
    Name ="Code"
    Expression ="[OLCC Items from OLAS Website].ItemID"
    Name ="Units Per Case"
    Expression ="[OLCC Items from OLAS Website].UnitsPerCase"
    Alias ="size"
    Name ="size"
    Expression ="Right([itemid],1)"
    Name ="Description"
    Expression ="[OLCC Items from OLAS Website].[Item Description]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="[OLCC Items from OLAS Website].ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[OLCC Items from OLAS Website].[Price_Per Case]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[OLCC Items from OLAS Website].UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[OLCC Items from OLAS Website].Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[OLCC Items from OLAS Website].[Item Description]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1537
    Bottom =877
    Left =-1
    Top =-1
    Right =1467
    Bottom =508
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =277
        Bottom =310
        Top =0
        Name ="OLCC Items from OLAS Website"
        Name =""
    End
End
