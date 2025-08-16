Operation =4
Option =0
Where ="(((PurchaseOrderLine.Cost)<>([PurchaseOrderLineUpdatePrice].[sprice])) AND ((Pur"
    "chaseOrderLineUpdatePrice.SalesDepartmentID)=1)) OR (((PurchaseOrderLine.Cost) I"
    "s Null) AND ((PurchaseOrderLineUpdatePrice.SalesDepartmentID)=1))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="PurchaseOrderLineUpdatePrice"
End
Begin OutputColumns
    Name ="PurchaseOrderLine.Cost"
    Expression ="[PurchaseOrderLineUpdatePrice].[sprice]"
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="PurchaseOrderLineUpdatePrice"
    Expression ="PurchaseOrderLine.ItemID = PurchaseOrderLineUpdatePrice.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="PurchaseOrderLine.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLineUpdatePrice.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-542
    Top =94
    Right =683
    Bottom =771
    Left =-1
    Top =-1
    Right =1193
    Bottom =269
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =28
        Top =41
        Right =201
        Bottom =155
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =373
        Top =30
        Right =517
        Bottom =291
        Top =0
        Name ="PurchaseOrderLineUpdatePrice"
        Name =""
    End
End
