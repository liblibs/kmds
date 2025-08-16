Operation =2
Name ="PurchaseOrderItemDaysSupply"
Option =0
Begin InputTables
    Name ="qryPurchaseOrderItemReplenish"
End
Begin OutputColumns
    Expression ="qryPurchaseOrderItemReplenish.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =162
        Top =6
        Right =530
        Bottom =120
        Top =0
        Name ="qryPurchaseOrderItemReplenish"
        Name =""
    End
End
