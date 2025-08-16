Operation =1
Option =0
Begin InputTables
    Name ="Sysfile"
    Name ="qryClubInvoiceForClubScreen"
End
Begin OutputColumns
    Expression ="qryClubInvoiceForClubScreen.LiquorItem"
    Expression ="Sysfile.*"
    Expression ="qryClubInvoiceForClubScreen.ClubID"
    Alias ="Order"
    Expression ="qryClubInvoiceForClubScreen.InvoiceNo"
    Expression ="qryClubInvoiceForClubScreen.InvoiceNo"
    Expression ="qryClubInvoiceForClubScreen.OrdDate"
    Expression ="qryClubInvoiceForClubScreen.InvDate"
    Expression ="qryClubInvoiceForClubScreen.LiquorExtPrice"
    Expression ="qryClubInvoiceForClubScreen.LiquorDiscount"
    Expression ="qryClubInvoiceForClubScreen.LiquorNetPrice"
    Expression ="qryClubInvoiceForClubScreen.NonLiquorExtPrice"
    Expression ="qryClubInvoiceForClubScreen.NonLiquorDiscount"
    Expression ="qryClubInvoiceForClubScreen.NonLiquorNetPrice"
    Alias ="ExtPrice"
    Expression ="[LiquorExtPrice]+[NonLiquorExtPrice]"
    Alias ="Discount"
    Expression ="[LiquorDiscount]+[NonLiquorDiscount]"
    Alias ="ExtNetPrice"
    Expression ="([LiquorExtPrice]+[NonLiquorExtPrice])-([LiquorDiscount]+[NonLiquorDiscount])"
    Expression ="qryClubInvoiceForClubScreen.NonLiquor"
    Expression ="qryClubInvoiceForClubScreen.ClubName"
    Expression ="qryClubInvoiceForClubScreen.ClubContact"
    Expression ="qryClubInvoiceForClubScreen.ClubAddress1"
    Expression ="qryClubInvoiceForClubScreen.ClubAddress2"
    Expression ="qryClubInvoiceForClubScreen.ClubCityST"
    Expression ="qryClubInvoiceForClubScreen.LineID"
    Expression ="qryClubInvoiceForClubScreen.ItemID"
    Expression ="qryClubInvoiceForClubScreen.Description"
    Expression ="qryClubInvoiceForClubScreen.SizeDesc"
    Expression ="qryClubInvoiceForClubScreen.QuantityShipped"
    Alias ="UNetPrice"
    Expression ="[Netprice]/[QuantityShipped]"
    Alias ="XPrice"
    Expression ="qryClubInvoiceForClubScreen.NetPrice"
    Alias ="XCost"
    Expression ="qryClubInvoiceForClubScreen.Cost"
    Alias ="UCost"
    Expression ="[Cost]/[QuantityShipped]"
    Alias ="OrderNotes"
    Expression ="qryClubInvoiceForClubScreen.ClubNotes"
    Expression ="qryClubInvoiceForClubScreen.SalePrice"
    Expression ="qryClubInvoiceForClubScreen.PricePerOZ"
    Expression ="qryClubInvoiceForClubScreen.ClubFax"
    Expression ="qryClubInvoiceForClubScreen.SalesDepartmentID"
    Expression ="qryClubInvoiceForClubScreen.CatDesc"
    Expression ="qryClubInvoiceForClubScreen.Location"
    Expression ="qryClubInvoiceForClubScreen.InvoiceNumber"
End
Begin OrderBy
    Expression ="qryClubInvoiceForClubScreen.LineID"
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
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="XCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderNotes"
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
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =475
        Bottom =255
        Top =0
        Name ="qryClubInvoiceForClubScreen"
        Name =""
    End
End
