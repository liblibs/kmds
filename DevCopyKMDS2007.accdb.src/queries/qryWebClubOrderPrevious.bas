Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\"))"
Having ="(((Club.ID)=[Forms]![ClubOrderWebImport]![cmbLicensee]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Club"
    Name ="Item"
    Name ="SubCategory"
End
Begin OutputColumns
    Alias ="ClubID"
    Expression ="Club.ID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="SubCategory.Description"
    Expression ="Item.Description"
    Expression ="InvoLine.Code"
    Alias ="TotalQty"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalCost"
    Expression ="Sum(InvoLine.Cost)"
    Alias ="Retail"
    Expression ="Sum(IIf([TransType]='C',[involine].[Price],0))"
    Alias ="Discount"
    Expression ="Sum(IIf([TransType]='D',[involine].[Price],0))"
    Alias ="TotalPrice"
    Expression ="Sum(InvoLine.Price)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept = Club.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SubCategory"
    Expression ="Item.SubCategoryID = SubCategory.ID"
    Flag =2
End
Begin OrderBy
    Expression ="SubCategory.Description"
    Flag =0
    Expression ="Item.Description"
    Flag =0
End
Begin Groups
    Expression ="Club.ID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="SubCategory.Description"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="InvoLine.Code"
    GroupLevel =0
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
        dbText "Name" ="TotalPrice"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Retail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4230"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SubCategory.Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =-457
    Top =420
    Right =1028
    Bottom =1097
    Left =-1
    Top =-1
    Right =1462
    Bottom =119
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =324
        Top =-3
        Right =555
        Bottom =157
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =603
        Top =12
        Right =781
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =829
        Top =12
        Right =973
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
