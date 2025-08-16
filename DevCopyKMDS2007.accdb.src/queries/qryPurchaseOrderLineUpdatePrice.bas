Operation =4
Option =0
Where ="(((PurchaseOrderLine.Price)<>[PurchaseOrderLineUpdatePrice].[sprice] Or (Purchas"
    "eOrderLine.Price) Is Null))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="PurchaseOrderLineUpdatePrice"
End
Begin OutputColumns
    Name ="PurchaseOrderLine.Price"
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
Begin
    Begin
        dbText "Name" ="PurchaseOrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLineUpdatePrice.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-541
    Top =128
    Right =684
    Bottom =805
    Left =-1
    Top =-1
    Right =1193
    Bottom =234
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =139
        Bottom =208
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =374
        Top =12
        Right =518
        Bottom =156
        Top =0
        Name ="PurchaseOrderLineUpdatePrice"
        Name =""
    End
End
