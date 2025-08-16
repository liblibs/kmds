Operation =1
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[parPurchaseOrderID]) AND ((PurchaseOrderL"
    "ine.QuantityShipped)<>0))"
Begin InputTables
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.VendorID"
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Expression ="PurchaseOrderLine.LineID"
    Expression ="PurchaseOrderLine.ItemID"
    Expression ="PurchaseOrderLine.QuantityShipped"
    Expression ="PurchaseOrderLine.Price"
    Expression ="PurchaseOrderLine.Cost"
    Expression ="PurchaseOrderLine.SetupDate"
    Expression ="PurchaseOrderLine.ChangeDate"
    Alias ="ExtPrice"
    Expression ="[QuantityOrdered]*[Price]"
    Alias ="ExtCost"
    Expression ="[QuantityOrdered]*[Cost]"
End
Begin Parameters
    Name ="parPurchaseOrderID"
    Flag =4
End
Begin OrderBy
    Expression ="PurchaseOrderLine.ItemID"
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
        dbText "Name" ="PurchaseOrderLine.ChangeDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.VendorID"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.PurchaseOrderID"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.LineID"
        dbInteger "ColumnWidth" ="720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.ItemID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
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
    Bottom =183
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =104
        Top =6
        Right =417
        Bottom =180
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
