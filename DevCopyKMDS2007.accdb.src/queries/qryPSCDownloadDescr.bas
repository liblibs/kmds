Operation =1
Option =0
Where ="(((Item.SalesDepartmentID) Between Forms!PSC!cboSalesDepartmentIDBegin And Forms"
    "!PSC!cboSalesDepartmentIDEnd) And ((Len(UPCItem.ID))=4 Or (Len(UPCItem.ID))=5 Or"
    " (Len(UPCItem.ID))=6 Or (Len(UPCItem.ID))=8 Or (Len(UPCItem.ID))=12))"
Begin InputTables
    Name ="UPCItem"
    Name ="Item"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="DESCR"
    Expression ="dscrptn(Item.Description) & \"  \" & Format([OnShelf],\"@@@\") & \" \" & Format("
        "Format(Item.price,\"#.00 \"),\"@@@@@@\") & Format([ItemID],\"@@@@@@\") & \" \" &"
        " Item.Description"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="Item.ID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
Begin OrderBy
    Expression ="UPCItem.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "DatasheetFontName" ="Courier"
dbInteger "DatasheetFontHeight" ="10"
dbInteger "DatasheetFontWeight" ="400"
dbBoolean "DatasheetFontItalic" ="0"
dbBoolean "DatasheetFontUnderline" ="0"
dbByte "TabularCharSet" ="0"
dbByte "TabularFamily" ="49"
dbLong "DatasheetForeColor" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbLong "DatasheetForeColor12" ="0"
Begin
    Begin
        dbText "Name" ="DESCR"
        dbInteger "ColumnWidth" ="6540"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
