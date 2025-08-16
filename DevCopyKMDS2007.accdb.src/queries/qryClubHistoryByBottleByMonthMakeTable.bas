Operation =2
Name ="ClubHistoryByBottleByMonth"
Option =0
Where ="(((InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\") And ((InvoLine.Setup"
    "Date) Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms!PrintInvento"
    "ryReport!txtDateRangeTo))"
Begin InputTables
    Name ="Club"
    Name ="qryClubInvoiceHistoryByInvoice"
    Name ="InvoLine"
    Name ="Item"
    Name ="Size"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="ClubID"
    Expression ="Club.ID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Alias ="CatDesc"
    Expression ="IIf(IsNull([subcategory].[Description]),\" Not Assigned\",[subcategory].[Descrip"
        "tion])"
    Expression ="InvoLine.Code"
    Alias ="Description"
    Expression ="IIf(IsNull(Item.Description),\"Deleted\",Item.Description)"
    Alias ="SizeD"
    Expression ="Size.Description"
    Alias ="Year-Month"
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
    Alias ="Qty"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="Dollars"
    Expression ="Sum(InvoLine.Price)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.SetupDate = InvoLine.SetupDate"
    Flag =1
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="qryClubInvoiceHistoryByInvoice"
    RightTable ="InvoLine"
    Expression ="qryClubInvoiceHistoryByInvoice.InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="Club"
    RightTable ="qryClubInvoiceHistoryByInvoice"
    Expression ="Club.ID = qryClubInvoiceHistoryByInvoice.ClubID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
End
Begin Groups
    Expression ="Club.ID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="IIf(IsNull([subcategory].[Description]),\" Not Assigned\",[subcategory].[Descrip"
        "tion])"
    GroupLevel =0
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="IIf(IsNull(Item.Description),\"Deleted\",Item.Description)"
    GroupLevel =0
    Expression ="Size.Description"
    GroupLevel =0
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Descr"
        dbInteger "ColumnWidth" ="3690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryClubSelect.ClubName"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="3690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.ClubName"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CatDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year-Month"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dollars"
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
    Bottom =149
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =74
        Top =34
        Right =170
        Bottom =148
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =268
        Top =20
        Right =517
        Bottom =177
        Top =0
        Name ="qryClubInvoiceHistoryByInvoice"
        Name =""
    End
    Begin
        Left =483
        Top =73
        Right =579
        Bottom =187
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =635
        Top =35
        Right =731
        Bottom =209
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =796
        Top =19
        Right =892
        Bottom =133
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =769
        Top =133
        Right =865
        Bottom =247
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
