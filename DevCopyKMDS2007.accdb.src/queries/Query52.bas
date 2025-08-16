Operation =1
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[PO]) AND ((OLCCEPS.OrderNumber)=[EPS]))"
Begin InputTables
    Name ="OLCCEPS"
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Expression ="OLCCEPS.OrderNumber"
    Expression ="OLCCEPS.OLCCItem"
    Expression ="OLCCEPS.Description"
    Expression ="OLCCEPS.Price"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="PurchaseOrderLine.QuantityOrdered"
End
Begin Joins
    LeftTable ="OLCCEPS"
    RightTable ="PurchaseOrderLine"
    Expression ="OLCCEPS.OLCCItem = PurchaseOrderLine.ItemID"
    Flag =2
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
        dbText "Name" ="PurchaseOrderLine.PurchaseOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OrderNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.OLCCItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCEPS.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-132
    Top =176
    Right =1100
    Bottom =934
    Left =-1
    Top =-1
    Right =1200
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =27
        Right =368
        Bottom =352
        Top =0
        Name ="OLCCEPS"
        Name =""
    End
    Begin
        Left =454
        Top =27
        Right =814
        Bottom =344
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
