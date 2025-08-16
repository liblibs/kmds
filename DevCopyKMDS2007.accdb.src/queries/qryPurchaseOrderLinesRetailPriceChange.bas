Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)<>1) And ((Item.Price)<>PurchaseOrderLine.price))"
Begin InputTables
    Name ="PurchaseOrderLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="PurchaseOrderLine.Price"
    Expression ="Item.Price"
End
Begin Joins
    LeftTable ="PurchaseOrderLine"
    RightTable ="Item"
    Expression ="PurchaseOrderLine.ItemID = Item.ID"
    Flag =1
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
    ColumnsShown =539
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
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
