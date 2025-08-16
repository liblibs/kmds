Operation =5
Option =0
Where ="(((UPCItem.SetupDate)<Now()-60) And ((InvoLine.InvoHeadID) Is Null) And ((Len(UP"
    "CItem.id))>Len([itemid])))"
Begin InputTables
    Name ="UPCItem"
    Name ="Item"
    Name ="InvoLine"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="UPCItem.SetupDate"
    Expression ="InvoLine.InvoHeadID"
    Alias ="Expr1"
    Expression ="Len(UPCItem.id)"
    Expression ="UPCItem.*"
    Alias ="Expr2"
    Expression ="Len([UPCItem].[id])"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="InvoLine"
    Expression ="UPCItem.ID = InvoLine.InputCode"
    Flag =2
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.SetupDate"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
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
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =135
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =267
        Top =138
        Right =363
        Bottom =252
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =268
        Top =14
        Right =364
        Bottom =128
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =498
        Top =114
        Right =594
        Bottom =228
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
