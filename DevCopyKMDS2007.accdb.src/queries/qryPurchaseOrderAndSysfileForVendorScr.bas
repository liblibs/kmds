Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qryItemTransactionInvoiceForVendors"
End
Begin OutputColumns
    Expression ="qryItemTransactionInvoiceForVendors.*"
    Expression ="Sysfile.AgentNumber"
    Expression ="Sysfile.Name"
    Expression ="Sysfile.Add1"
    Expression ="Sysfile.Add2"
    Expression ="Sysfile.Add3"
    Expression ="Sysfile.Adv3"
    Expression ="Sysfile.Phone"
    Expression ="Sysfile.BlankLines"
    Expression ="Sysfile.CurrShiftDate"
    Expression ="Sysfile.SalesTaxRate"
    Expression ="Sysfile.Cr1Installed"
    Expression ="Sysfile.Cr2Installed"
    Expression ="Sysfile.Cr3Installed"
    Expression ="Sysfile.Cr4Installed"
    Expression ="Sysfile.Cr1Online"
    Expression ="Sysfile.Cr2Online"
    Expression ="Sysfile.Cr3Online"
    Expression ="Sysfile.Cr4Online"
    Expression ="Sysfile.Cr1CheckOut"
    Expression ="Sysfile.Cr2CheckOut"
    Expression ="Sysfile.Cr3CheckOut"
    Expression ="Sysfile.Cr4CheckOut"
    Expression ="Sysfile.Online"
    Expression ="Sysfile.CashDropItemID"
    Expression ="Sysfile.CashLoanItemID"
    Expression ="Sysfile.ChargeDropItemID"
    Expression ="Sysfile.CheckDropItemID"
    Expression ="Sysfile.ClubInvItemID"
    Expression ="Sysfile.ClubDiscRate"
    Expression ="Sysfile.ClubDiscntItemID"
    Expression ="Sysfile.CouponItemID"
    Expression ="Sysfile.DiscountItemID"
    Expression ="Sysfile.LotteryPaidOutItemID"
    Expression ="Sysfile.NonClubDiscntItemID"
    Expression ="Sysfile.NonClubDiscRate"
    Expression ="Sysfile.NonClubInvItemID"
    Expression ="Sysfile.CancelTenderID"
    Expression ="Sysfile.NoSaleTenderID"
    Expression ="Sysfile.ScannedODLTenderID"
    Expression ="Sysfile.ReportsDir"
    Expression ="Sysfile.CashsysInsLoc"
    Expression ="Sysfile.NetworkDatabase"
    Expression ="Sysfile.NetworkMailDir"
    Expression ="Sysfile.OLCCVendor"
    Expression ="Sysfile.InvoiceCounter"
    Expression ="Sysfile.SetupDate"
    Expression ="Sysfile.ChangeDate"
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
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =17
        Top =19
        Right =113
        Bottom =133
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =161
        Top =79
        Right =526
        Bottom =193
        Top =0
        Name ="qryItemTransactionInvoiceForVendors"
        Name =""
    End
End
