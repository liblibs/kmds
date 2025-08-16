Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\"))"
Having ="(((Club.ID)=[Forms]![Clubs]![ID]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Club"
End
Begin OutputColumns
    Alias ="ClubID"
    Expression ="Club.ID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="InvoLine.InvoHeadID"
    Alias ="TotalQty"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="PostDate"
    Expression ="InvoLine.SetupDate"
    Alias ="TotalCost"
    Expression ="Sum(InvoLine.Cost)"
    Expression ="InvoLine.CashSysID"
    Alias ="Retail"
    Expression ="Sum(IIf([TransType]='C',[Price],0))"
    Alias ="Discount"
    Expression ="Sum(IIf([TransType]='D',[Price],0))"
    Alias ="TotalPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.SalesGroup"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept=Club.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =1
End
Begin Groups
    Expression ="Club.ID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.SetupDate"
    GroupLevel =0
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.SalesGroup"
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
        dbText "Name" ="TotalPrice"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Retail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PostDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
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
    Bottom =215
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
End
