Operation =1
Option =0
Where ="(((Club.ID)=[Forms]![ClubOrderPrintDialog]![ID]))"
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="Club"
    Name ="ClubOrderGuides"
    Name ="qryItemsSalePriceNow"
End
Begin OutputColumns
    Expression ="Club.ID"
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
    Expression ="ClubOrderGuides.LineID"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="Size.Oz"
    Alias ="PricePerShot"
    Expression ="IIf([SalesDepartmentID]<>1,[Price],CCur((([price]/Size.oz)*club.OuncesColumn)*0."
        "95))"
    Expression ="qryItemsSalePriceNow.SalePrice"
    Expression ="Item.Price"
    Alias ="TodaysPrice"
    Expression ="IIf(IsNull([SalePrice]),[Price],[SalePrice])"
    Alias ="Sprice"
    Expression ="IIf(IsNull([SalePrice]),\"\",\"*\")"
    Alias ="ClubFax"
    Expression ="Club.Phone2"
    Expression ="Club.MaintainParLevels"
    Expression ="ClubOrderGuides.Par"
    Expression ="Item.Location"
    Expression ="Item.SalesDepartmentID"
    Expression ="Club.OuncesColumn"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID=Size.ID"
    Flag =2
    LeftTable ="Club"
    RightTable ="ClubOrderGuides"
    Expression ="Club.ID=ClubOrderGuides.ClubID"
    Flag =1
    LeftTable ="Item"
    RightTable ="ClubOrderGuides"
    Expression ="Item.ID=ClubOrderGuides.ItemID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID=qryItemsSalePriceNow.ItemID"
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
        dbText "Name" ="TodaysPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sprice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricePerShot"
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
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =286
        Top =56
        Right =382
        Bottom =170
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =471
        Top =185
        Right =567
        Bottom =299
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =7
        Top =4
        Right =103
        Bottom =208
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =143
        Top =31
        Right =239
        Bottom =145
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
    Begin
        Left =463
        Top =24
        Right =559
        Bottom =138
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
End
