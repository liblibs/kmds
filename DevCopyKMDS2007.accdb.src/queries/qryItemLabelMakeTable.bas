Operation =2
Name ="ItemLabels"
Option =0
Begin InputTables
    Name ="Item"
    Name ="SalesDepartment"
    Name ="qryItemsSalePriceNow"
    Name ="Size"
    Name ="Vendor"
    Name ="qryUPCItemMax"
End
Begin OutputColumns
    Expression ="Item.ID"
    Alias ="ItemD"
    Expression ="Item.Description"
    Alias ="SizeD"
    Expression ="IIf(size.Description=\"Not Assigned\",\"\",size.description)"
    Alias ="CurPrice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Alias ="S"
    Expression ="IIf(IsNull([SalePrice]),\"\",\"*\")"
    Alias ="SalesDpt"
    Expression ="SalesDepartment.Description"
    Expression ="Item.Price"
    Alias ="VendorName"
    Expression ="Vendor.Name"
    Expression ="qryUPCItemMax.UPCCodeMax"
    Expression ="Item.SetupDate"
    Expression ="Item.ChangeDate"
    Alias ="Labels"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryUPCItemMax"
    Expression ="Item.ID = qryUPCItemMax.ItemID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="qryItemLabel.UPCCodeMax"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="CurPrice"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDpt"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S"
        dbInteger "ColumnWidth" ="285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorName"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsSalePriceNow.SalePrice"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurSale"
        dbInteger "ColumnWidth" ="630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemD"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Labels"
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
    Bottom =510
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =44
        Top =41
        Right =140
        Bottom =245
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =341
        Top =129
        Right =437
        Bottom =243
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =814
        Top =7
        Right =965
        Bottom =121
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =652
        Top =74
        Right =748
        Bottom =188
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =206
        Top =170
        Right =302
        Bottom =284
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =762
        Top =169
        Right =858
        Bottom =268
        Top =0
        Name ="qryUPCItemMax"
        Name =""
    End
End
