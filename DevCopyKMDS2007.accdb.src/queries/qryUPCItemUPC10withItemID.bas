Operation =1
Option =0
Where ="(((Len(upcitem.id))=12) And ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="UPCItem"
    Name ="Item"
End
Begin OutputColumns
    Expression ="UPCItem.ID"
    Expression ="UPCItem.ItemID"
    Alias ="lid"
    Expression ="Len([upcitem].[id])"
    Expression ="Item.Description"
    Expression ="Item.SalesDepartmentID"
    Alias ="upc10"
    Expression ="Mid([UPCITEM].[ID],2,10)"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Len([upcitem].[id])"
    Flag =0
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
        dbText "Name" ="UPCItem.ID"
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
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upc10"
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
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =194
        Top =0
        Name ="Item"
        Name =""
    End
End
