Operation =4
Option =0
Where ="(((Item.PriceGroupID)<>0) AND ((PriceGroup.ID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="PriceGroup"
End
Begin OutputColumns
    Name ="Item.PriceGroupID"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="PriceGroup"
    Expression ="Item.PriceGroupID = PriceGroup.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.PLUList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.DateLastSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.AverageDailySales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Deposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.PriceGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceGroup.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SubPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.TransferredToday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.PictureID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.OrderAt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.OrderTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.AvgTurns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPackPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.DaysSupplyGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesLeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.RestockAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.TPR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1266
    Bottom =846
    Left =-1
    Top =-1
    Right =1234
    Bottom =522
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =255
        Bottom =429
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =379
        Top =72
        Right =616
        Bottom =399
        Top =0
        Name ="PriceGroup"
        Name =""
    End
End
