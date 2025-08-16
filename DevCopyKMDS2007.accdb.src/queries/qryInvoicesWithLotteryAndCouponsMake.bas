Operation =2
Name ="InvoLineWithLotteryAndCoupons"
Option =0
Begin InputTables
    Name ="qryInvoicesWithCoupons"
    Name ="qryInvoicesWithLottery"
End
Begin OutputColumns
    Expression ="qryInvoicesWithCoupons.InvoHeadID"
    Expression ="qryInvoicesWithCoupons.CashSysID"
    Expression ="qryInvoicesWithCoupons.SetupDate"
End
Begin Joins
    LeftTable ="qryInvoicesWithCoupons"
    RightTable ="qryInvoicesWithLottery"
    Expression ="qryInvoicesWithCoupons.InvoHeadID = qryInvoicesWithLottery.InvoHeadID"
    Flag =1
    LeftTable ="qryInvoicesWithCoupons"
    RightTable ="qryInvoicesWithLottery"
    Expression ="qryInvoicesWithCoupons.CashSysID = qryInvoicesWithLottery.CashSysID"
    Flag =1
    LeftTable ="qryInvoicesWithCoupons"
    RightTable ="qryInvoicesWithLottery"
    Expression ="qryInvoicesWithCoupons.SetupDate = qryInvoicesWithLottery.SetupDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =146
        Bottom =105
        Top =0
        Name ="qryInvoicesWithCoupons"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =294
        Bottom =105
        Top =0
        Name ="qryInvoicesWithLottery"
        Name =""
    End
End
