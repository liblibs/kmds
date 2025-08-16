Operation =4
Option =0
Where ="(((Len([txtQuantity]))=12) AND ((MetrologicScanbatch.txtQuantity)<\"000000001000"
    "\"))"
Begin InputTables
    Name ="MetrologicScanbatch"
End
Begin OutputColumns
    Name ="MetrologicScanbatch.Quantity"
    Expression ="Mid([txtQuantity],9,3)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="MetrologicScanbatch.txtQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.Quantity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-468
    Top =429
    Right =1099
    Bottom =1194
    Left =-1
    Top =-1
    Right =1535
    Bottom =366
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =282
        Top =0
        Name ="MetrologicScanbatch"
        Name =""
    End
End
