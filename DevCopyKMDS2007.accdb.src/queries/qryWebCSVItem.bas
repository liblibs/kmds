Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.SizeID"
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.VendorID"
    Expression ="Item.SubCategoryID"
    Expression ="vwItemsWithCurrentPricing.CurPrice"
    Alias ="OnSale"
    Expression ="IIf([SPrice]<>0,True,False)"
    Expression ="Item.Cost"
    Expression ="Item.Onhand"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="Item.ChangeDate"
    Expression ="Item.Description"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="Item.ID = vwItemsWithCurrentPricing.ID"
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
dbMemo "Filter" ="([qryWebCSVItem].[ID]=60152)"
Begin
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
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ChangeDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.CurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vwItemsWithCurrentPricing.OnShelf"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-10
    Top =159
    Right =1514
    Bottom =822
    Left =-1
    Top =-1
    Right =1492
    Bottom =283
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =261
        Bottom =469
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =485
        Top =9
        Right =719
        Bottom =375
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
