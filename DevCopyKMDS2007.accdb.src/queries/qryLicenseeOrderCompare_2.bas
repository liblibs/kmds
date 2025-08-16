Operation =1
Option =0
Where ="(((OrderLine.OrdheadID)=[Forms]![ClubOrderCompare]![txtPurchaseOrderID_2]))"
Begin InputTables
    Name ="Item"
    Name ="OrderLine"
End
Begin OutputColumns
    Expression ="OrderLine.OrdheadID"
    Expression ="OrderLine.ItemID"
    Expression ="Item.Description"
    Expression ="OrderLine.QuantityShipped"
    Expression ="OrderLine.Price"
End
Begin Joins
    LeftTable ="OrderLine"
    RightTable ="Item"
    Expression ="OrderLine.ItemID = Item.ID"
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
    Begin
        dbText "Name" ="OrderLine.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Price"
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
    Bottom =285
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =335
        Top =12
        Right =649
        Bottom =261
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =697
        Top =12
        Right =841
        Bottom =266
        Top =0
        Name ="OrderLine"
        Name =""
    End
End
