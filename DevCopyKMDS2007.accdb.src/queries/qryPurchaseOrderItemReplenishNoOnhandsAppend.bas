Operation =3
Name ="PurchaseOrderItemDaysSupply"
Option =0
Where ="(((Item.VendorID)=[parVendorID]) AND ((Item.OrderTo)>0) AND ((PurchaseOrderItemD"
    "aysSupply.ItemId) Is Null) AND ((Item.Onhand)<=0))"
Begin InputTables
    Name ="Item"
    Name ="PurchaseOrderItemDaysSupply"
End
Begin OutputColumns
    Name ="VendorID"
    Expression ="Item.VendorID"
    Name ="Description"
    Expression ="Item.Description"
    Name ="ItemId"
    Expression ="Item.ID"
    Alias ="SOQ"
    Name ="SOQ"
    Expression ="Item.OrderTo"
    Name ="SPrice"
    Expression ="Item.Price"
    Name ="Cost"
    Expression ="Item.Cost"
    Alias ="Repack"
    Name ="Repack"
    Expression ="0"
End
Begin Parameters
    Name ="parVendorID"
    Flag =4
End
Begin Joins
    LeftTable ="Item"
    RightTable ="PurchaseOrderItemDaysSupply"
    Expression ="Item.ID = PurchaseOrderItemDaysSupply.ItemId"
    Flag =2
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
    Flag =0
    Expression ="Item.Description"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="SOQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
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
    Bottom =375
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =375
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =527
        Bottom =120
        Top =0
        Name ="PurchaseOrderItemDaysSupply"
        Name =""
    End
End
