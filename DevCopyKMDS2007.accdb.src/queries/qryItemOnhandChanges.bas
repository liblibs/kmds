Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryAuditReport!cboSalesDepartmen"
    "tIDBegin And Forms!PrintInventoryAuditReport!cboSalesDepartmentIDEnd) And ((Cash"
    "ier.ID) Between Forms!PrintInventoryAuditReport!cboCashiersIDBegin And Forms!Pri"
    "ntInventoryAuditReport!cboCashiersIDEnd) And ((InvoLine.SetupDate) Between Forms"
    "!PrintInventoryAuditReport!dtpDateBegin And Forms!PrintInventoryAuditReport!dtpD"
    "ateEnd) And ((InvoLine.TransType)=\"A\" Or (InvoLine.TransType)=\"R\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="Cashier"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Alias ="SD"
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
    Expression ="SalesDepartment.SalesGroupID"
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Alias ="UPrice"
    Expression ="IIf([Quantity]=0,0,InvoLine.Price/[Quantity])"
    Alias ="Ext"
    Expression ="InvoLine.Price"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InputCode"
    Alias ="NewOnShelf"
    Expression ="InvoLine.SalesGroup"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.SalesGroupID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Item.ID"
    Flag =0
    Expression ="Cashier.ID"
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
        dbText "Name" ="Item.ID"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.ID"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.SalesGroupID"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ext"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewOnShelf"
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
    Bottom =395
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =352
        Top =44
        Right =513
        Bottom =263
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =589
        Top =82
        Right =685
        Bottom =196
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =196
        Top =45
        Right =308
        Bottom =324
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =28
        Top =170
        Right =149
        Bottom =374
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
