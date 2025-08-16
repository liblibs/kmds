Operation =1
Option =0
Where ="(((Item.SalesDepartmentID) Between [Forms]![PrintInventoryReport]![cboSalesDepar"
    "tmentIDBegin] And [Forms]![PrintInventoryReport]![cboSalesDepartmentIDEnd]))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
    Name ="Size"
    Name ="qryItemsSalePriceNow"
    Name ="qryItemOnReserve"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Alias ="DepartmentDesc"
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeD"
    Expression ="Size.Description"
    Expression ="Item.Onhand"
    Expression ="Item.Cost"
    Expression ="Item.Price"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Alias ="TodaysPrice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Expression ="qryItemOnReserve.OnReserve"
    Alias ="OnTheShelf"
    Expression ="[onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTo"
        "dayCr4]"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
    Flag =2
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
    Flag =0
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
        dbText "Name" ="TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnTheShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =179
        Top =12
        Right =345
        Bottom =261
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =397
        Top =176
        Right =493
        Bottom =290
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =460
        Top =2
        Right =650
        Bottom =116
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
    Begin
        Left =536
        Top =130
        Right =632
        Bottom =214
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
End
