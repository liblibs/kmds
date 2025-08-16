Operation =1
Option =0
Begin InputTables
    Name ="KitItem"
End
Begin OutputColumns
    Expression ="KitItem.KitID"
    Alias ="Retail"
    Expression ="Sum(KitItem.Price)"
End
Begin Groups
    Expression ="KitItem.KitID"
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
        dbText "Name" ="KitItem.KitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Retail"
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
    Bottom =337
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =221
        Top =0
        Name ="KitItem"
        Name =""
    End
End
