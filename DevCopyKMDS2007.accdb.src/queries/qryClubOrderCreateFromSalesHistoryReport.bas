Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"C\"))"
Having ="(((Club.ID)=[Forms]![ClubOrderCreate]![lstClubs]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="MinOfSetupDate"
    Expression ="Min(InvoLine.SetupDate)"
    Alias ="MaxOfSetupDate"
    Expression ="Max(InvoLine.SetupDate)"
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="Item.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="LastOfPrice"
    Expression ="Last(Item.Price)"
    Alias ="ExtPrice"
    Expression ="Sum(InvoLine.price)"
    Alias ="LastOfCost"
    Expression ="Last(Item.Cost)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept = Club.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
    Expression ="Min(InvoLine.SetupDate)"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="Club.ID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="InvoLine.SalesDept"
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
        dbText "Name" ="InvoLine.Code"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.ID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="LastOfPrice"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfCost"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfSetupDate"
        dbInteger "ColumnWidth" ="2250"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfSetupDate"
        dbInteger "ColumnWidth" ="2250"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbInteger "ColumnOrder" ="6"
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
    Bottom =217
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =350
        Bottom =286
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =780
        Top =0
        Right =958
        Bottom =246
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =548
        Top =150
        Right =692
        Bottom =294
        Top =0
        Name ="Club"
        Name =""
    End
End
