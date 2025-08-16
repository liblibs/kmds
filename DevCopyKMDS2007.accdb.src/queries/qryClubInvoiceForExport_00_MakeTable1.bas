Operation =2
Name ="ClubInvoiceForCSVExport"
Option =0
Begin InputTables
    Name ="qryClubInvoice"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="Sysfile.AgentNumber"
    Expression ="Sysfile.Name"
    Expression ="qryClubInvoice.ClubID"
    Expression ="qryClubInvoice.ClubName"
    Expression ="qryClubInvoice.InvoiceNo"
    Expression ="qryClubInvoice.InvDate"
    Expression ="qryClubInvoice.ItemID"
    Expression ="qryClubInvoice.Description"
    Alias ="Size"
    Expression ="qryClubInvoice.SizeDesc"
    Alias ="Quantity"
    Expression ="qryClubInvoice.QuantityShipped"
    Alias ="Price"
    Expression ="qryClubInvoice.NetPrice"
    Alias ="XPrice"
    Expression ="[QuantityShipped]*[NetPrice]"
    Expression ="qryClubInvoice.SalePrice"
    Expression ="qryClubInvoice.PricePerOZ"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="XPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.PricePerOZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.OuncesColumn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.InvDate"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-448
    Top =0
    Right =1257
    Bottom =856
    Left =-1
    Top =-1
    Right =1673
    Bottom =138
    Left =0
    Top =8
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =163
        Bottom =270
        Top =0
        Name ="qryClubInvoice"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =378
        Bottom =222
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
