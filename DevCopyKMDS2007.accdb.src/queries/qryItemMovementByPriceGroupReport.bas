Operation =1
Option =0
Where ="(((PriceGroup.Description) Between Forms!PrintInventoryReport!cboPriceGroupIDBeg"
    "in And Forms!PrintInventoryReport!cboPriceGroupIDEnd) And ((InvoLine.SetupDate) "
    "Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms!PrintInventoryRepo"
    "rt!txtDateRangeTo))"
Begin InputTables
    Name ="PriceGroup"
    Name ="InvoLine"
    Name ="Item"
    Name ="TranTypes"
    Name ="Cashier"
End
Begin OutputColumns
    Alias ="PriceGrouipID"
    Expression ="PriceGroup.ID"
    Alias ="PriceGroupDesc"
    Expression ="PriceGroup.Description"
    Alias ="ItemID"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="InvoLine.SetupDate"
    Alias ="Trans"
    Expression ="TranTypes.Description"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
    Alias ="Purchases"
    Expression ="IIf(\"P\"=TranTypes.Code,[Quantity],0)"
    Alias ="Counter"
    Expression ="IIf(\" Sales - Counter\"=[Trans],[Quantity],0)"
    Alias ="Clubs"
    Expression ="IIf(\" Sales - Clubs\"=[Trans],[Quantity],0)"
    Expression ="InvoLine.Quantity"
    Alias ="Transfers"
    Expression ="IIf(\"X\"=TranTypes.Code,[Quantity],0)"
    Alias ="Adjustments"
    Expression ="IIf(\"A\"=TranTypes.Code,[Quantity],0)"
    Alias ="Emp"
    Expression ="Cashier.Name"
    Alias ="TranType"
    Expression ="TranTypes.Code"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType = TranTypes.Code"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =2
    LeftTable ="PriceGroup"
    RightTable ="Item"
    Expression ="PriceGroup.ID = Item.PriceGroupID"
    Flag =1
End
Begin OrderBy
    Expression ="PriceGroup.ID"
    Flag =0
    Expression ="Item.Description"
    Flag =0
    Expression ="InvoLine.SetupDate"
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trans"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Purchases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transfers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjustments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceGrouipID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceGroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Emp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranType"
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
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="PriceGroup"
        Name =""
    End
    Begin
        Left =416
        Top =-7
        Right =512
        Bottom =182
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =190
        Top =12
        Right =317
        Bottom =171
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =690
        Top =97
        Right =786
        Bottom =196
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =605
        Top =-5
        Right =701
        Bottom =109
        Top =0
        Name ="Cashier"
        Name =""
    End
End
