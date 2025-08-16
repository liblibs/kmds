Operation =3
Name ="ClubInvoiceForCSVExport"
Option =0
Begin InputTables
    Name ="qryClubInvoice"
    Name ="Sysfile"
End
Begin OutputColumns
    Alias ="FirstOfAgentNumber"
    Name ="AgentNumber"
    Expression ="First(Sysfile.AgentNumber)"
    Alias ="FirstOfName"
    Name ="Name"
    Expression ="First(Sysfile.Name)"
    Alias ="FirstOfClubID"
    Name ="ClubID"
    Expression ="First(qryClubInvoice.ClubID)"
    Alias ="FirstOfClubName"
    Name ="ClubName"
    Expression ="First(qryClubInvoice.ClubName)"
    Alias ="FirstOfInvoiceNo"
    Name ="InvoiceNo"
    Expression ="First(qryClubInvoice.InvoiceNo)"
    Alias ="FirstOfInvDate"
    Name ="InvDate"
    Expression ="First(qryClubInvoice.InvDate)"
    Alias ="FirstOfClubDiscntItemID"
    Name ="ItemID"
    Expression ="First(Sysfile.ClubDiscntItemID)"
    Alias ="des"
    Name ="Description"
    Expression ="First(\"LiquorDiscount\")"
    Alias ="Size"
    Name ="Size"
    Expression ="First(\"LiqDisc\")"
    Alias ="Quantity"
    Name ="Quantity"
    Expression ="First(1)"
    Alias ="Price"
    Name ="Price"
    Expression ="First(0-[LiquorDiscount])"
    Alias ="XPrice"
    Name ="XPrice"
    Expression ="First(0-[LiquorDiscount])"
    Alias ="FirstOfSalePrice"
    Name ="SalePrice"
    Expression ="First(qryClubInvoice.SalePrice)"
    Alias ="PricePerOZ"
    Name ="PricePerOZ"
    Expression ="First(0)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="PricePerOZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OuncesColumn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfAgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfInvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfInvDate"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfClubDiscntItemID"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="Sysfile.ClubDiscntItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="des"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubInvoice.LiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfSalePrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-874
    Top =131
    Right =831
    Bottom =987
    Left =-1
    Top =-1
    Right =1673
    Bottom =155
    Left =0
    Top =8
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =274
        Bottom =270
        Top =0
        Name ="qryClubInvoice"
        Name =""
    End
    Begin
        Left =350
        Top =10
        Right =556
        Bottom =226
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
