Operation =4
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[parPurchaseOrderID]) AND ((PurchaseOrderL"
    "ine.QuantityShipped)<>0))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Onhand"
    Expression ="[Onhand]+[QuantityShipped]"
    Name ="Item.Cost"
    Expression ="PurchaseOrderLine.cost"
End
Begin Parameters
    Name ="parPurchaseOrderID"
    Flag =4
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="Item"
    Expression ="PurchaseOrderLine.ItemID = Item.ID"
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
    Bottom =329
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =222
        Bottom =270
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
    Begin
        Left =260
        Top =6
        Right =406
        Bottom =300
        Top =0
        Name ="Item"
        Name =""
    End
End
