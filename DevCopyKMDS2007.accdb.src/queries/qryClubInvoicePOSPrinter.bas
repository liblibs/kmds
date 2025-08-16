Operation =1
Option =0
Where ="(((OrdHead.InvoiceNo)=[parInvoiceNo]) AND ((OrderLine.QuantityShipped)<>0))"
Begin InputTables
    Name ="Club"
    Name ="OrdHead"
    Name ="Item"
    Name ="Size"
    Name ="OrderLine"
End
Begin OutputColumns
    Expression ="OrdHead.ID"
    Expression ="OrdHead.ClubID"
    Expression ="OrdHead.Order"
    Expression ="OrdHead.InvoiceNo"
    Expression ="OrdHead.OrdDate"
    Expression ="OrdHead.InvDate"
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
    Expression ="Size.Oz"
    Alias ="PricePerOZ"
    Expression ="CCur(OrderLine.price/[oz])"
End
Begin Parameters
    Name ="parInvoiceNo"
    Flag =4
End
Begin Joins
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
        Left =25
        Top =133
        Right =143
        Bottom =277
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
        Left =512
        Top =8
        Right =608
        Bottom =182
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
End
