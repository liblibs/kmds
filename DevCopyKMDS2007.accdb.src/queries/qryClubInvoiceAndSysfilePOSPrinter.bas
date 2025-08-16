Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qryClubInvoicePOSPrinter"
End
Begin OutputColumns
    Expression ="Sysfile.*"
    Expression ="qryClubInvoicePOSPrinter.ClubID"
    Expression ="qryClubInvoicePOSPrinter.Order"
    Expression ="qryClubInvoicePOSPrinter.InvoiceNo"
    Expression ="qryClubInvoicePOSPrinter.OrdDate"
    Expression ="qryClubInvoicePOSPrinter.InvDate"
    Expression ="qryClubInvoicePOSPrinter.ClubName"
    Expression ="qryClubInvoicePOSPrinter.ClubContact"
    Expression ="qryClubInvoicePOSPrinter.ClubAddress1"
    Expression ="qryClubInvoicePOSPrinter.ClubAddress2"
    Expression ="qryClubInvoicePOSPrinter.ClubCityST"
    Expression ="qryClubInvoicePOSPrinter.ClubPhone1"
    Expression ="qryClubInvoicePOSPrinter.LineID"
    Expression ="qryClubInvoicePOSPrinter.ItemID"
    Expression ="qryClubInvoicePOSPrinter.Description"
    Expression ="qryClubInvoicePOSPrinter.SizeDesc"
    Expression ="qryClubInvoicePOSPrinter.QuantityShipped"
    Expression ="qryClubInvoicePOSPrinter.Price"
    Alias ="XPrice"
    Expression ="[QuantityShipped]*[Price]"
    Alias ="XCost"
    Expression ="[QuantityShipped]*[Cost]"
    Expression ="qryClubInvoicePOSPrinter.Cost"
    Expression ="qryClubInvoicePOSPrinter.OrderNotes"
    Expression ="qryClubInvoicePOSPrinter.SalePrice"
    Expression ="qryClubInvoicePOSPrinter.PricePerOZ"
End
Begin OrderBy
    Expression ="qryClubInvoicePOSPrinter.ItemID"
    Flag =0
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
        dbText "Name" ="XPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XCost"
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
    Bottom =274
    Left =0
    Top =8
    ColumnsShown =539
    Begin
        Left =358
        Top =-3
        Right =454
        Bottom =111
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =492
        Top =6
        Right =728
        Bottom =232
        Top =0
        Name ="qryClubInvoicePOSPrinter"
        Name =""
    End
End
