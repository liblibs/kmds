Operation =3
Name ="OrdHead"
Option =0
Begin InputTables
    Name ="qryFixClubOrders"
End
Begin OutputColumns
    Name ="ClubID"
    Expression ="qryFixClubOrders.ClubID"
    Name ="Order"
    Expression ="qryFixClubOrders.InvoiceNumber"
    Name ="InvoiceNo"
    Expression ="qryFixClubOrders.InvoiceNumber"
    Name ="SetupDate"
    Expression ="qryFixClubOrders.SetupDate"
    Name ="InvDate"
    Expression ="qryFixClubOrders.SetupDate"
    Name ="OrdDate"
    Expression ="qryFixClubOrders.SetupDate"
    Name ="ChangeDate"
    Expression ="qryFixClubOrders.SetupDate"
    Alias ="p"
    Name ="DisDollars"
    Expression ="0-[Price]"
    Alias ="q"
    Name ="LiquorDiscount"
    Expression ="0-[Price]"
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
dbBoolean "UseTransaction" ="-1"
Begin
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
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFixClubOrders.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFixClubOrders.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-103
    Top =9
    Right =1042
    Bottom =666
    Left =-1
    Top =-1
    Right =1113
    Bottom =340
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =260
        Top =11
        Right =496
        Bottom =328
        Top =0
        Name ="qryFixClubOrders"
        Name =""
    End
End
