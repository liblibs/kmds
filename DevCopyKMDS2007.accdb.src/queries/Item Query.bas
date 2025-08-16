Operation =1
Option =0
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.[ID]"
    Expression ="Item.[SizeID]"
    Expression ="Item.[SalesDepartmentID]"
    Expression ="Item.[VendorID]"
    Expression ="Item.[PriceGroupID]"
    Expression ="Item.[SubCategoryID]"
    Expression ="Item.[Description]"
    Expression ="Item.[Onhand]"
    Expression ="Item.[OrderAt]"
    Expression ="Item.[OrderTo]"
    Expression ="Item.[AvgTurns]"
    Expression ="Item.[Price]"
    Expression ="Item.[Cost]"
    Expression ="Item.[Deposit]"
    Expression ="Item.[UnitsPerCase]"
    Expression ="Item.[UnitsPerOrder]"
    Expression ="Item.[UnitsPerPack]"
    Expression ="Item.[SubPack]"
    Expression ="Item.[SoldTodayCr1]"
    Expression ="Item.[SoldTodayCr2]"
    Expression ="Item.[SoldTodayCr3]"
    Expression ="Item.[SoldTodayCr4]"
    Expression ="Item.[TransferredToday]"
    Expression ="Item.[RestockAmount]"
    Expression ="Item.[Location]"
    Expression ="Item.[QtyPack]"
    Expression ="Item.[QtyPackPrice]"
    Expression ="Item.[StockNumber]"
    Expression ="Item.[PLUList]"
    Expression ="Item.[DateLastSold]"
    Expression ="Item.[TPR]"
    Expression ="Item.[AverageDailySales]"
    Expression ="Item.[DaysSupplyGroupID]"
    Expression ="Item.[SetupDate]"
    Expression ="Item.[ChangeDate]"
    Expression ="Item.[upsize_ts]"
    Expression ="Item.[SoldTodayCr5]"
    Expression ="Item.[SalesLeadID]"
    Expression ="Item.[PictureID]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Item.[ID]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =136
    Right =1297
    Bottom =882
    Left =-1
    Top =-1
    Right =1187
    Bottom =408
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
End
