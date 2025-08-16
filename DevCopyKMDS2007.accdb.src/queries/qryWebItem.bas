Operation =1
Option =0
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.SizeID"
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.VendorID"
    Expression ="Item.PriceGroupID"
    Expression ="Item.SubCategoryID"
    Expression ="Item.Description"
    Expression ="Item.Onhand"
    Expression ="Item.OrderAt"
    Expression ="Item.OrderTo"
    Expression ="Item.AvgTurns"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="Item.Deposit"
    Expression ="Item.UnitsPerCase"
    Expression ="Item.UnitsPerOrder"
    Expression ="Item.UnitsPerPack"
    Expression ="Item.SubPack"
    Expression ="Item.SoldTodayCr1"
    Expression ="Item.SoldTodayCr2"
    Expression ="Item.SoldTodayCr3"
    Expression ="Item.SoldTodayCr4"
    Expression ="Item.TransferredToday"
    Expression ="Item.RestockAmount"
    Expression ="Item.Location"
    Expression ="Item.QtyPack"
    Expression ="Item.QtyPackPrice"
    Expression ="Item.StockNumber"
    Expression ="Item.PLUList"
    Expression ="Item.DateLastSold"
    Expression ="Item.TPR"
    Expression ="Item.AverageDailySales"
    Expression ="Item.DaysSupplyGroupID"
    Expression ="Item.SetupDate"
    Expression ="Item.ChangeDate"
    Expression ="Item.upsize_ts"
    Expression ="Item.SoldTodayCr5"
    Expression ="Item.SalesLeadID"
    Expression ="Item.PictureID"
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
        dbText "Name" ="Item.PictureID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.PriceGroupID"
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
        dbText "Name" ="Item.ID"
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
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerPack"
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
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
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
        dbText "Name" ="Item.SubCategoryID"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1131
    Bottom =804
    Left =-1
    Top =-1
    Right =1099
    Bottom =464
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =350
        Top =0
        Name ="Item"
        Name =""
    End
End
