Operation =1
Option =0
Begin InputTables
    Name ="PSCInput"
    Name ="UPCItem"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(PSCInput.Quantity)"
End
Begin Joins
    LeftTable ="PSCInput"
    RightTable ="UPCItem"
    Expression ="PSCInput.Code = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="UPCItem.ItemID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin Groups
    Expression ="vwItemsWithCurrentPricing.ID"
    GroupLevel =0
    Expression ="vwItemsWithCurrentPricing.Description"
    GroupLevel =0
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
        dbText "Name" ="vwItemsWithCurrentPricing.Description"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =90
        Top =0
        Name ="PSCInput"
        Name =""
    End
    Begin
        Left =272
        Top =-3
        Right =368
        Bottom =111
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =506
        Top =-5
        Right =602
        Bottom =109
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
