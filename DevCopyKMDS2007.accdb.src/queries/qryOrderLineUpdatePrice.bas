Operation =4
Option =0
Where ="(((OrderLine.Price)<>qryPurchaseOrderItemsWithTodaysPrice.sprice Or (OrderLine.P"
    "rice) Is Null) And ((qryPurchaseOrderItemsWithTodaysPrice.SalesDepartmentID)=1))"
Begin InputTables
    Name ="OrderLine"
    Name ="qryPurchaseOrderItemsWithTodaysPrice"
End
Begin OutputColumns
    Name ="OrderLine.Price"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.sprice"
    Name ="OrderLine.SalePrice"
    Expression ="qryPurchaseOrderItemsWithTodaysPrice.Star"
End
Begin Joins
    LeftTable ="OrderLine"
    RightTable ="qryPurchaseOrderItemsWithTodaysPrice"
    Expression ="OrderLine.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID"
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
    Bottom =281
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =36
        Top =45
        Right =167
        Bottom =279
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =245
        Top =56
        Right =620
        Bottom =260
        Top =0
        Name ="qryPurchaseOrderItemsWithTodaysPrice"
        Name =""
    End
End
