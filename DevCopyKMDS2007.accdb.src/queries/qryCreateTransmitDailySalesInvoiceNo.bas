Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="InvoLine5"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.SetupDate"
    Expression ="InvoLine5.CashSysID"
    Alias ="DiscountID"
    Expression ="InvoLine5.Code"
    Alias ="Discount"
    Expression ="InvoLine5.Price"
    Alias ="LicenseeInvoiceNumber"
    Expression ="Format([InputCode],\"0000000000\")"
    Alias ="ClubID"
    Expression ="Club.ID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Alias ="PremisesNumber"
    Expression ="Format([Licensee],\"0000000000\")"
End
Begin Joins
    LeftTable ="Sysfile"
    RightTable ="InvoLine5"
    Expression ="Sysfile.ClubDiscntItemID = InvoLine5.Code"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Club"
    Expression ="InvoLine5.SalesDept = Club.ID"
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
        dbText "Name" ="ClubName"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LicenseeInvoiceNumber"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PremisesNumber"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DiscountID"
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
        Left =38
        Top =6
        Right =219
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =302
        Top =7
        Right =398
        Bottom =121
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =436
        Top =6
        Right =532
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
End
