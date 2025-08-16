Operation =2
Name ="DaysSupplyADSForNewItems"
Option =0
Where ="(((qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate)>Now()-365) AND ((q"
    "ryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate)>Now()-90))"
Begin InputTables
    Name ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
End
Begin OutputColumns
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfQuantity"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfPrice"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate"
    Expression ="qryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate"
    Alias ="Days"
    Expression ="CInt([MaxOfSetupDate]-[MinOfSetupDate])"
    Alias ="ADS"
    Expression ="([SumOfQuantity]/IIf([Days]<30,30,[days]))"
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
    Top =0
    ColumnsShown =539
    Begin
        Left =59
        Top =-2
        Right =155
        Bottom =112
        Top =0
        Name ="qryDaysSupplyInvoLineItemSalesByMinMaxDate"
        Name =""
    End
End
