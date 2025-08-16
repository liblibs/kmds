Operation =1
Option =0
Having ="(((Count(UPCItem.ItemID))=1))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ItemID"
    Alias ="CountOfItemID"
    Expression ="Count(UPCItem.ItemID)"
End
Begin Groups
    Expression ="UPCItem.ItemID"
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
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1506
    Bottom =876
    Left =-1
    Top =-1
    Right =1474
    Bottom =552
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
