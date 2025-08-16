Operation =1
Option =0
Begin InputTables
    Name ="qryClubInvoice"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="qryClubInvoice.LiquorItem"
    Expression ="Sysfile.*"
    Expression ="qryClubInvoice.ClubID"
    Expression ="qryClubInvoice.Order"
    Expression ="qryClubInvoice.InvoiceNo"
    Expression ="qryClubInvoice.OrdDate"
    Expression ="qryClubInvoice.InvDate"
    Expression ="qryClubInvoice.LiquorExtPrice"
    Expression ="qryClubInvoice.LiquorDiscount"
    Expression ="qryClubInvoice.LiquorNetPrice"
    Expression ="qryClubInvoice.NonLiquorExtPrice"
    Expression ="qryClubInvoice.NonLiquorDiscount"
    Expression ="qryClubInvoice.NonLiquorNetPrice"
    Alias ="ExtPrice"
    Expression ="[LiquorExtPrice]+[NonLiquorExtPrice]"
    Alias ="Discount"
    Expression ="[LiquorDiscount]+[NonLiquorDiscount]"
    Alias ="ExtNetPrice"
    Expression ="([LiquorExtPrice]+[NonLiquorExtPrice])-([LiquorDiscount]+[NonLiquorDiscount])"
    Expression ="qryClubInvoice.NonLiquor"
    Expression ="qryClubInvoice.ClubName"
    Expression ="qryClubInvoice.ClubContact"
    Expression ="qryClubInvoice.ClubAddress1"
    Expression ="qryClubInvoice.ClubAddress2"
    Expression ="qryClubInvoice.ClubCityST"
    Expression ="qryClubInvoice.ClubPhone1"
    Expression ="qryClubInvoice.LineID"
    Expression ="qryClubInvoice.ItemID"
    Expression ="qryClubInvoice.Description"
    Expression ="qryClubInvoice.SizeDesc"
    Expression ="qryClubInvoice.QuantityShipped"
    Expression ="qryClubInvoice.NetPrice"
    Alias ="XPrice"
    Expression ="[QuantityShipped]*[NetPrice]"
    Alias ="XCost"
    Expression ="[QuantityShipped]*[Cost]"
    Expression ="qryClubInvoice.Cost"
    Expression ="qryClubInvoice.OrderNotes"
    Expression ="qryClubInvoice.SalePrice"
    Expression ="qryClubInvoice.PricePerOZ"
    Expression ="qryClubInvoice.ClubFax"
    Expression ="qryClubInvoice.SalesDepartmentID"
    Expression ="qryClubInvoice.CatDesc"
    Expression ="qryClubInvoice.Location"
    Expression ="qryClubInvoice.OuncesColumn"
    Expression ="qryClubInvoice.DiscountedPrice"
End
Begin OrderBy
    Expression ="qryClubInvoice.LineID"
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
        Right =268
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
