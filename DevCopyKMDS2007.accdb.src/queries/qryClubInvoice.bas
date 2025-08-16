Operation =1
Option =0
Where ="(((OrdHead.InvoiceNo)=Forms!ClubOrderPrintDialog!txtInvoiceNo) And ((OrderLine.Q"
    "uantityShipped)<>0))"
Begin InputTables
    Name ="Club"
    Name ="OrdHead"
    Name ="qryClubOrderLineTotals"
    Name ="Item"
    Name ="Size"
    Name ="OrderLine"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="LiquorItem"
    Expression ="IIf([SalesDepartmentID]=1,-1,0)"
    Expression ="OrdHead.ID"
    Expression ="OrdHead.ClubID"
    Expression ="OrdHead.Order"
    Expression ="OrdHead.InvoiceNo"
    Expression ="OrdHead.OrdDate"
    Expression ="OrdHead.InvDate"
    Expression ="qryClubOrderLineTotals.Quantity"
    Expression ="qryClubOrderLineTotals.ExtCost"
    Expression ="qryClubOrderLineTotals.LiquorExtPrice"
    Expression ="qryClubOrderLineTotals.LiquorDiscount"
    Alias ="LiquorNetPrice"
    Expression ="[LiquorExtPrice]-[LiquorDiscount]"
    Expression ="qryClubOrderLineTotals.NonLiquorExtPrice"
    Expression ="qryClubOrderLineTotals.NonLiquorDiscount"
    Alias ="NonLiquorNetPrice"
    Expression ="[NonLiquorExtPrice]-[NonLiquorDiscount]"
    Expression ="qryClubOrderLineTotals.NonLiquor"
    Alias ="ClubName"
    Expression ="Club.Name"
    Alias ="ClubContact"
    Expression ="Club.Contact"
    Alias ="ClubAddress1"
    Expression ="Club.Address1"
    Alias ="ClubAddress2"
    Expression ="Club.Address2"
    Alias ="ClubCityST"
    Expression ="[City] & \", \" & [ST] & \"  \" & [Zip]"
    Alias ="ClubPhone1"
    Expression ="Club.Phone1"
    Expression ="OrderLine.LineID"
    Expression ="OrderLine.ItemID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="OrderLine.QuantityShipped"
    Expression ="OrderLine.Price"
    Expression ="OrderLine.Cost"
    Expression ="OrdHead.OrderNotes"
    Expression ="OrderLine.SalePrice"
    Expression ="OrderLine.Buydown"
    Alias ="NetPrice"
    Expression ="OrderLine.Price-[BuyDown]"
    Alias ="Ounces"
    Expression ="IIf((IsNull(size.Oz) Or size.oz=0),1,size.oz)"
    Alias ="PricePerOZ"
    Expression ="IIf(item.salesdepartmentId<>1,orderline.price,CCur(((OrderLine.price/[Ounces])*C"
        "lub.OuncesColumn)*0.95))"
    Alias ="ClubFax"
    Expression ="Club.Phone2"
    Expression ="Item.SalesDepartmentID"
    Alias ="CatDesc"
    Expression ="SubCategory.Description"
    Expression ="Item.Location"
    Expression ="Club.OuncesColumn"
    Alias ="DiscountedPrice"
    Expression ="IIf(item.salesdepartmentId<>1,(orderline.price-[buydown]),CCur(OrderLine.price*0"
        ".95))"
End
Begin Joins
    LeftTable ="OrdHead"
    RightTable ="qryClubOrderLineTotals"
    Expression ="OrdHead.ID = qryClubOrderLineTotals.OrdheadID"
    Flag =1
    LeftTable ="Club"
    RightTable ="OrdHead"
    Expression ="Club.ID = OrdHead.ClubID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="OrderLine"
    Expression ="Item.ID = OrderLine.ItemID"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="OrderLine"
    Expression ="OrdHead.ID = OrderLine.OrdheadID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
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
        dbText "Name" ="ClubCityST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricePerOZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ounces"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DiscountedPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubContact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubPhone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CatDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.LiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.NonLiquorExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.OrderNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.OuncesColumn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.OrdDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.LiquorExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.NonLiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubOrderLineTotals.NonLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderLine.Buydown"
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
    Bottom =241
    Left =0
    Top =8
    ColumnsShown =539
    Begin
        Left =20
        Top =141
        Right =138
        Bottom =285
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =207
        Top =39
        Right =303
        Bottom =153
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =7
        Top =14
        Right =160
        Bottom =128
        Top =0
        Name ="qryClubOrderLineTotals"
        Name =""
    End
    Begin
        Left =512
        Top =8
        Right =608
        Bottom =272
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =646
        Top =16
        Right =742
        Bottom =160
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =355
        Top =28
        Right =451
        Bottom =217
        Top =0
        Name ="OrderLine"
        Name =""
    End
    Begin
        Left =666
        Top =152
        Right =762
        Bottom =266
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
