Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Name ="PurchaseOrderLine.Repack"
    Expression ="IIf([QuantityOrdered]-(([quantityOrdered]\\[UnitsPerCase])*[UnitsPerCase])=0,0,-"
        "1)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="PurchaseOrderLine"
    Expression ="Item.ID = PurchaseOrderLine.ItemID"
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
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =172
        Top =6
        Right =325
        Bottom =240
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =210
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
