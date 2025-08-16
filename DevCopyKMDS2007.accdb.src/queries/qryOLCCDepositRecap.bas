Operation =1
Option =0
Begin InputTables
    Name ="qryOLCCDepositRecapSales"
    Name ="qryOLCCDepositRecapTenders"
End
Begin OutputColumns
    Expression ="qryOLCCDepositRecapSales.Counter"
    Expression ="qryOLCCDepositRecapSales.ClubsGross"
    Expression ="qryOLCCDepositRecapSales.ClubsDiscount"
    Expression ="qryOLCCDepositRecapSales.ClubsNet"
    Expression ="qryOLCCDepositRecapSales.OLCCDeposit"
    Expression ="qryOLCCDepositRecapTenders.Checks"
    Expression ="qryOLCCDepositRecapTenders.Cards"
    Alias ="CashRequired"
    Expression ="[OLCCDeposit]-[Checks]-[Cards]-[EFT]"
    Alias ="Deposit"
    Expression ="[OLCCDeposit]-[Cards]-[EFT]"
    Expression ="qryOLCCDepositRecapTenders.EFT"
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
        dbText "Name" ="CashRequired"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapSales.Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapTenders.EFT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapSales.ClubsGross"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapSales.ClubsDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapSales.ClubsNet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapSales.OLCCDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapTenders.Checks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCDepositRecapTenders.Cards"
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
    Right =1193
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =174
        Bottom =120
        Top =0
        Name ="qryOLCCDepositRecapSales"
        Name =""
    End
    Begin
        Left =365
        Top =-10
        Right =629
        Bottom =94
        Top =0
        Name ="qryOLCCDepositRecapTenders"
        Name =""
    End
End
