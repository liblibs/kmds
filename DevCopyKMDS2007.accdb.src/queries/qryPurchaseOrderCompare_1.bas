Operation =1
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[Forms]![PurchaseOrderCompare]![txtPurchas"
    "eOrderID_1]))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="Item.Description"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Expression ="PurchaseOrderLine.Price"
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="Item"
    Expression ="PurchaseOrderLine.ItemID=Item.ID"
    Flag =1
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
        dbText "Name" ="PurchaseOrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =38
    Right =983
    Bottom =601
    Left =-1
    Top =-1
    Right =907
    Bottom =262
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =287
        Bottom =262
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =335
        Top =12
        Right =566
        Bottom =261
        Top =0
        Name ="Item"
        Name =""
    End
End
