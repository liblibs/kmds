Operation =2
Name ="NewUPC"
Option =0
Where ="(((Item.SalesDepartmentID)>2))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="UPCItem.ID"
    Alias ="newUPC"
    Expression ="upcatoe([upcitem].[id])"
    Expression ="UPCItem.ItemID"
    Alias ="lid"
    Expression ="Len(upcatoe([upcitem].[id]))"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="newUPC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1497
    Bottom =877
    Left =-1
    Top =-1
    Right =1465
    Bottom =554
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
