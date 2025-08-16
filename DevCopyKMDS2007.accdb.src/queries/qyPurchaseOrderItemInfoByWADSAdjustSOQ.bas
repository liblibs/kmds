Operation =4
Option =0
Where ="(((PurchaseOrderItemWADS.SOQ)>0 And (PurchaseOrderItemWADS.SOQ)<([SOQRaw]-0.01))"
    ")"
Begin InputTables
    Name ="PurchaseOrderItemWADS"
End
Begin OutputColumns
    Name ="PurchaseOrderItemWADS.SOQ"
    Expression ="[SOQ]+[OrderMultiple]"
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
    Bottom =192
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =204
        Top =0
        Name ="PurchaseOrderItemWADS"
        Name =""
    End
End
