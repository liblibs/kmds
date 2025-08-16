Operation =1
Option =0
Where ="(((InvoLine5.ID)=1))"
Begin InputTables
    Name ="qryFinTechTransactions"
    Name ="InvoLine5"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.ID"
    Expression ="InvoLine5.CashSysID"
    Alias ="ClubID"
    Expression ="InvoLine5.SalesDept"
    Expression ="InvoLine5.SetupDate"
    Expression ="Club.Name"
    Expression ="Club.Licensee"
    Expression ="Club.FintechAccount"
    Alias ="Invoice_Number"
    Expression ="CStr([SalesGroup])"
End
Begin Joins
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.CashSysID = InvoLine5.CashSysID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.SetupDate = InvoLine5.SetupDate"
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
        dbText "Name" ="InvoLine5.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Licensee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.FintechAccount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice_Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-235
    Top =105
    Right =1347
    Bottom =898
    Left =-1
    Top =-1
    Right =1559
    Bottom =475
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =397
        Top =0
        Name ="qryFinTechTransactions"
        Name =""
    End
    Begin
        Left =379
        Top =9
        Right =523
        Bottom =463
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =725
        Top =19
        Right =869
        Bottom =422
        Top =0
        Name ="Club"
        Name =""
    End
End
