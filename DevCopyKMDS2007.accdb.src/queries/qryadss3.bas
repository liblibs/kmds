Operation =3
Name ="DaysSupplyADSForNewItems"
Option =0
Where ="(((qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate)>Now()-365) AND ((q"
    "ryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate)>Now()-90) AND ((DaysSu"
    "pplyADSForNewItems.Code) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
    Name ="DaysSupplyADSForNewItems"
End
Begin OutputColumns
    Name ="Code"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code"
    Name ="SumOfQuantity"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfQuantity"
    Name ="SumOfPrice"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfPrice"
    Name ="MinOfSetupDate"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate"
    Name ="MaxOfSetupDate"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate"
    Alias ="Days"
    Name ="Days"
    Expression ="CInt([MaxOfSetupDate]-[MinOfSetupDate])"
    Alias ="ADS"
    Name ="ADS"
    Expression ="(([SumOfQuantity]/IIf([Days]<30,30,[days]))/[UnitsPerPack])"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.SubPack = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
    Expression ="UPCItem.ItemID = qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code"
    Flag =1
    LeftTable ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
    RightTable ="DaysSupplyADSForNewItems"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code = DaysSupplyADSForNewItems.Code"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="qryInvoLineItemSalesByMinMaxDate.MinOfSetupDate"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="days"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =110
    Left =0
    Top =7
    ColumnsShown =651
    Begin
        Left =817
        Top =-10
        Right =962
        Bottom =104
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =649
        Top =2
        Right =745
        Bottom =116
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =190
        Top =-3
        Right =609
        Bottom =111
        Top =0
        Name ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="DaysSupplyADSForNewItems"
        Name =""
    End
End
