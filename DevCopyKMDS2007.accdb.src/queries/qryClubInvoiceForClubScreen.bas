Operation =1
Option =0
Where ="(((InvoLine.CashSysID)=Forms!Clubs!ClubsPreviousOrderSub.Form!CashSysID) And ((I"
    "nvoLine.InvoHeadID)=Forms!Clubs!ClubsPreviousOrderSub.Form!InvoHeadID) And ((Inv"
    "oLine.SetupDate)=Forms!Clubs!ClubsPreviousOrderSub.Form!PostDate) And ((InvoLine"
    ".Quantity)<>0 And (InvoLine.Quantity)<>0))"
Begin InputTables
    Name ="Club"
    Name ="InvoLine"
    Name ="Item"
    Name ="Size"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="LiquorItem"
    Expression ="IIf([SalesDepartmentID]=1,-1,0)"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
    Alias ="ClubID"
    Expression ="Club.ID"
    Alias ="InvoiceNo"
    Expression ="InvoLine.InvoHeadID"
    Alias ="OrdDate"
    Expression ="InvoLine.SetupDate"
    Alias ="InvDate"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.Quantity"
    Alias ="ExtCost"
    Expression ="InvoLine.Cost"
    Alias ="LiquorExtPrice"
    Expression ="1"
    Alias ="LiquorDiscount"
    Expression ="1"
    Alias ="LiquorNetPrice"
    Expression ="[LiquorExtPrice]-[LiquorDiscount]"
    Alias ="NonLiquorExtPrice"
    Expression ="1"
    Alias ="NonLiquorDiscount"
    Expression ="1"
    Alias ="NonLiquorNetPrice"
    Expression ="[NonLiquorExtPrice]-[NonLiquorDiscount]"
    Expression ="Club.NonLiquor"
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
    Alias ="LineID"
    Expression ="InvoLine.ID"
    Alias ="ItemID"
    Expression ="InvoLine.Code"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Alias ="QuantityShipped"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.Cost"
    Expression ="Club.ClubNotes"
    Alias ="SalePrice"
    Expression ="\"\""
    Alias ="Buydown"
    Expression ="\"\""
    Alias ="NetPrice"
    Expression ="InvoLine.Price"
    Alias ="Ounces"
    Expression ="IIf((IsNull(size.Oz) Or size.oz=0),1,size.oz)"
    Alias ="PricePerOZ"
    Expression ="CCur(IIf([involine].[transtype]=\"D\",0,([InvoLine].[price]/[InvoLine].[Quantity"
        "])/[Ounces]))"
    Alias ="ClubFax"
    Expression ="Club.Phone2"
    Expression ="Item.SalesDepartmentID"
    Alias ="CatDesc"
    Expression ="SubCategory.Description"
    Expression ="Item.Location"
    Alias ="InvoiceNumber"
    Expression ="InvoLine.SalesGroup"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Club"
    RightTable ="InvoLine"
    Expression ="Club.ID = InvoLine.SalesDept"
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
        dbText "Name" ="LiquorItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorNetPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubCityST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buydown"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ounces"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricePerOZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
        dbInteger "ColumnWidth" ="750"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
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
        dbText "Name" ="LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NetPrice"
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
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.ClubNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.NonLiquor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1275
    Bottom =854
    Left =-1
    Top =-1
    Right =1251
    Bottom =428
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
        Left =349
        Top =45
        Right =481
        Bottom =349
        Top =0
        Name ="InvoLine"
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
        Left =648
        Top =171
        Right =744
        Bottom =285
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
