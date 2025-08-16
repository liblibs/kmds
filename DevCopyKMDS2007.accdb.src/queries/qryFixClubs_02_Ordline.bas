Operation =3
Name ="OrderLine"
Option =0
Where ="(((InvoLine.TransType)=\"C\"))"
Begin InputTables
    Name ="qryFixClubOrders"
    Name ="InvoLine"
    Name ="OrdHead"
End
Begin OutputColumns
    Alias ="OrdheadID"
    Name ="OrdheadID"
    Expression ="OrdHead.ID"
    Alias ="LineID"
    Name ="LineID"
    Expression ="InvoLine.ID"
    Alias ="ItemID"
    Name ="ItemID"
    Expression ="InvoLine.Code"
    Name ="QuantityOrdered"
    Expression ="InvoLine.Quantity"
    Name ="QuantityShipped"
    Expression ="InvoLine.Quantity"
    Alias ="UP"
    Name ="Price"
    Expression ="[InvoLine].[Price]/[InvoLine].[Quantity]"
    Alias ="UT"
    Name ="Cost"
    Expression ="[InvoLine].[Price]/[InvoLine].[Quantity]"
    Name ="SetupDate"
    Expression ="InvoLine.SetupDate"
    Name ="ChangeDate"
    Expression ="InvoLine.SetupDate"
End
Begin Joins
    LeftTable ="qryFixClubOrders"
    RightTable ="InvoLine"
    Expression ="qryFixClubOrders.InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="qryFixClubOrders"
    RightTable ="InvoLine"
    Expression ="qryFixClubOrders.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="qryFixClubOrders"
    RightTable ="InvoLine"
    Expression ="qryFixClubOrders.SetupDate = InvoLine.SetupDate"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="qryFixClubOrders"
    Expression ="OrdHead.InvoiceNo = qryFixClubOrders.InvoiceNumber"
    Flag =1
End
Begin OrderBy
    Expression ="OrdHead.ID"
    Flag =0
    Expression ="InvoLine.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFixClubOrders.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdheadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.OrdDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFixClubOrders.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFixClubOrders.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1007"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UT"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-222
    Top =44
    Right =923
    Bottom =701
    Left =-1
    Top =-1
    Right =1113
    Bottom =323
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =293
        Top =36
        Right =529
        Bottom =353
        Top =0
        Name ="qryFixClubOrders"
        Name =""
    End
    Begin
        Left =643
        Top =16
        Right =787
        Bottom =362
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =44
        Top =20
        Right =220
        Bottom =342
        Top =0
        Name ="OrdHead"
        Name =""
    End
End
