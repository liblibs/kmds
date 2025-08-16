Operation =4
Option =0
Where ="(((PurchaseOrderLine.Price)<>qryPurchaseOrderItemsWithTodaysPrice.sprice Or (Pur"
    "chaseOrderLine.Price) Is Null)) Or (((PurchaseOrderLine.Cost)<>qryPurchaseOrderI"
    "temsWithTodaysPrice.Cost Or (PurchaseOrderLine.Cost) Is Null))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="qryPurchaseOrderItemsWithTodaysPrice"
End
Begin OutputColumns
    Name ="PurchaseOrderLine.Price"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.sprice"
    Name ="PurchaseOrderLine.Cost"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.cost"
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="qryPurchaseOrderItemsWithTodaysPrice"
    Expression ="PurchaseOrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =615
        Bottom =120
        Top =0
        Name ="qryPurchaseOrderItemsWithTodaysPrice"
        Name =""
    End
End
