Operation =1
Option =0
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SizeID"
    Alias ="count"
    Expression ="Sum(1)"
End
Begin Groups
    Expression ="Item.SizeID"
    GroupLevel =0
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
        dbText "Name" ="count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
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
    Bottom =353
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
End
