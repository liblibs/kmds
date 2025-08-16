Operation =1
Option =0
Begin InputTables
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Alias ="Quantity"
    Expression ="Sum(PurchaseOrderLine.QuantityShipped)"
    Alias ="Dollars"
    Expression ="Sum([QuantityShipped]*[Price])"
End
Begin Groups
    Expression ="PurchaseOrderLine.PurchaseOrderID"
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
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dollars"
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
    Bottom =249
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =179
        Bottom =241
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
