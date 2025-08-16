Operation =3
Name ="ItemsSoldToday"
Option =0
Begin InputTables
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Name ="Code"
    Expression ="Item.ID"
    Alias ="Qty"
    Name ="SoldToday"
    Expression ="Sum(IIf([TransType]=\"x\",0-[Quantity],[Quantity]))"
    Alias ="SumOfPrice"
    Name ="DollarsSoldToday"
    Expression ="Sum(InvoLine5.Price)"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="Item.ID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="SumOfQuantity"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =169
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =180
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =424
        Top =23
        Right =520
        Bottom =137
        Top =0
        Name ="Item"
        Name =""
    End
End
