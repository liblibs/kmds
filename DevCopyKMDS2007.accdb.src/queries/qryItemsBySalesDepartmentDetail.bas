Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)>2) AND ((Item.Price)>0) AND ((Item.PriceGroupID)=0))"
Begin InputTables
    Name ="Item"
    Name ="SubCategory"
    Name ="SalesDepartment"
    Name ="Vendor"
End
Begin OutputColumns
    Alias ="SalesDeptID"
    Expression ="Item.SalesDepartmentID"
    Alias ="SalesDept"
    Expression ="SalesDepartment.Description"
    Expression ="Item.VendorID"
    Expression ="Item.Description"
    Expression ="Item.ID"
    Alias ="Category"
    Expression ="SubCategory.Description"
    Expression ="Item.Price"
    Expression ="Item.Cost"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID=SubCategory.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID=SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID=Vendor.ID"
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
    Begin
        dbText "Name" ="SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDeptID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =161
    Top =38
    Right =1065
    Bottom =864
    Left =-1
    Top =-1
    Right =872
    Bottom =343
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =180
        Bottom =300
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =228
        Top =227
        Right =372
        Bottom =371
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =241
        Top =22
        Right =385
        Bottom =129
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =436
        Top =114
        Right =580
        Bottom =258
        Top =0
        Name ="Vendor"
        Name =""
    End
End
