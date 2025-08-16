Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"C\"))"
Having ="(((InvoLine.SalesDept)=[parClubID]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="FirstOfSetupDate"
    Expression ="First(InvoLine.SetupDate)"
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="Item.Description"
    Alias ="LastOfPrice"
    Expression ="Last(Item.Price)"
    Alias ="LastOfCost"
    Expression ="Last(Item.Cost)"
    Alias ="LastOfSetupDate"
    Expression ="Last(InvoLine.SetupDate)"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
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
    Expression ="First(InvoLine.SetupDate)"
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FirstOfSetupDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="LastOfSetupDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfPrice"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfCost"
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
    Bottom =336
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
        Left =421
        Top =12
        Right =599
        Bottom =258
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =789
        Top =94
        Right =933
        Bottom =238
        Top =0
        Name ="Club"
        Name =""
    End
End
