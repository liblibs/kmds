Operation =4
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[parPurchaseOrderID]) AND ((PurchaseOrderL"
    "ine.PrintLabels)<>0))"
Begin InputTables
    Name ="ItemLabels"
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Name ="ItemLabels.Labels"
    Expression ="IIf([PurchaseOrderLine].[QuantityShipped]<>0,[PurchaseOrderLine].[QuantityShippe"
        "d],1)"
End
Begin Parameters
    Name ="parPurchaseOrderID"
    Flag =4
End
Begin Joins
    LeftTable ="ItemLabels"
    RightTable ="PurchaseOrderLine"
    Expression ="ItemLabels.ID = PurchaseOrderLine.ItemID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemLabels.Labels"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="ItemLabels"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =353
        Bottom =120
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
