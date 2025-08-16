Operation =3
Name ="OLCCItemsDetailList"
Option =0
Begin InputTables
    Name ="OLCCItemsDetailListXLS"
End
Begin OutputColumns
    Name ="OLCCCode"
    Expression ="OLCCItemsDetailListXLS.F1"
    Name ="Desription"
    Expression ="OLCCItemsDetailListXLS.F2"
    Name ="Category"
    Expression ="OLCCItemsDetailListXLS.F3"
    Name ="Size"
    Expression ="OLCCItemsDetailListXLS.F4"
    Name ="Status"
    Expression ="OLCCItemsDetailListXLS.F5"
    Name ="UnitsPerOrder"
    Expression ="OLCCItemsDetailListXLS.F6"
    Name ="UnitsPerCase"
    Expression ="OLCCItemsDetailListXLS.F7"
    Name ="PricePerBottle"
    Expression ="OLCCItemsDetailListXLS.F8"
    Name ="QuantityAvailable"
    Expression ="OLCCItemsDetailListXLS.F9"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItemsDetailListXLS.F8"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =120
    Right =1486
    Bottom =873
    Left =-1
    Top =-1
    Right =1421
    Bottom =467
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =138
        Top =53
        Right =282
        Bottom =446
        Top =0
        Name ="OLCCItemsDetailListXLS"
        Name =""
    End
End
