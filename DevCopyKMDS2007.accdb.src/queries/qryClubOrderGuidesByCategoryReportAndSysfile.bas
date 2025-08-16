Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qryClubOrderGuidesReportByCategory"
End
Begin OutputColumns
    Alias ="STRIPPED"
    Expression ="IIf(([ItemID]>99900000000),[ItemID]-99900000000,[ItemID])"
    Expression ="qryClubOrderGuidesReportByCategory.*"
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
    Alias ="LineDesc"
    Expression ="Format(qryClubOrderGuidesReportByCategory.LineID,\"#\") & \" - \" & qryClubOrder"
        "GuidesReportByCategory.Description"
    Expression ="qryClubOrderGuidesReportByCategory.Par"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="LineDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Par"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr1Installed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr1Online"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr2Online"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr3Online"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr4Online"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr1CheckOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Online"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CashDropItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ChargeDropItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ClubInvItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CouponItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ReportsDir"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NetworkMailDir"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="STRIPPED"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubContact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubPhone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Add3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr4Installed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr4CheckOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.LotteryPaidOutItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NonClubInvItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NoSaleTenderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.InvoiceCounter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubCityST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Adv3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.SalesTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Size.Oz"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.qryItemsSalePriceNow.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Sprice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubOrderGuides.Par"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Add1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr3Installed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr3CheckOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CashLoanItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CheckDropItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ClubDiscntItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NonClubDiscntItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NetworkDatabase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.OLCCVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Phone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CurrShiftDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.PricePerOZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Add2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr2Installed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr2CheckOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ClubDiscRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.DiscountItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.NonClubDiscRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CancelTenderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ScannedODLTenderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.CashsysInsLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Club.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubOrderGuides.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.ClubOrderGuides.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderGuidesReportByCategory.Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.BlankLines"
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
    Bottom =119
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="qryClubOrderGuidesReportByCategory"
        Name =""
    End
End
