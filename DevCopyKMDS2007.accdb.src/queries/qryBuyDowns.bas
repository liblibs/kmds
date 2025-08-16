Operation =1
Option =0
Begin InputTables
    Name ="BuyDown"
    Name ="BuyDownItem"
    Name ="Item"
    Name ="Item"
    Alias ="Item_1"
    Name ="Size"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="BuyDown.ID"
    Expression ="BuyDown.RepsName"
    Expression ="BuyDown.ReferenceNo"
    Expression ="BuyDown.BeginDate"
    Expression ="BuyDown.EndDate"
    Expression ="BuyDown.BuydownDollars"
    Expression ="BuyDown.BuydownQty"
    Alias ="PromoItemID"
    Expression ="BuyDown.ItemsID"
    Alias ="PromItemDesc"
    Expression ="Item.Description"
    Expression ="BuyDownItem.ItemsCode"
    Expression ="Item_1.Description"
    Expression ="BuyDownItem.QuantitySold"
    Expression ="BuyDownItem.DiscountDollars"
    Expression ="BuyDownItem.ShpQty"
    Expression ="BuyDownItem.ShpDollars"
    Alias ="SalesdptD"
    Expression ="SalesDepartment.Description"
    Alias ="SizeD"
    Expression ="Size.Description"
End
Begin Joins
    LeftTable ="BuyDown"
    RightTable ="BuyDownItem"
    Expression ="BuyDown.ID = BuyDownItem.BuyDownID"
    Flag =1
    LeftTable ="BuyDown"
    RightTable ="Item"
    Expression ="BuyDown.ItemsID = Item.ID"
    Flag =1
    LeftTable ="BuyDownItem"
    RightTable ="Item_1"
    Expression ="BuyDownItem.ItemsCode = Item_1.ID"
    Flag =1
    LeftTable ="Item_1"
    RightTable ="Size"
    Expression ="Item_1.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item_1"
    RightTable ="SalesDepartment"
    Expression ="Item_1.SalesDepartmentID = SalesDepartment.ID"
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
        dbText "Name" ="BuyDown.BuydownDollars"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PromoItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PromItemDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesdptD"
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
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =163
        Top =37
        Right =300
        Bottom =151
        Top =0
        Name ="BuyDown"
        Name =""
    End
    Begin
        Left =329
        Top =33
        Right =564
        Bottom =282
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
    Begin
        Left =22
        Top =-6
        Right =118
        Bottom =108
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =629
        Top =6
        Right =725
        Bottom =210
        Top =0
        Name ="Item_1"
        Name =""
    End
    Begin
        Left =763
        Top =6
        Right =859
        Bottom =120
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =774
        Top =127
        Right =870
        Bottom =241
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
