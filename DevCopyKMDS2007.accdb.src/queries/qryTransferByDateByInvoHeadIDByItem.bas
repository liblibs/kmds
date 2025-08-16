Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And F"
    "orms!PrintInventoryReport!txtDateRangeTo) And ((InvoLine.TransType)=\"X\") And ("
    "(Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentID"
    "Begin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd Or (Item.SalesDepar"
    "tmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Forms!"
    "PrintInventoryReport!cboSalesDepartmentIDEnd) And ((InvoLine.Drawer)=0))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.ShiftID"
    Expression ="InvoLine.InputCode"
    Alias ="LiquorQuantity"
    Expression ="IIf(Item.SalesDepartmentID=1,[Quantity],0)"
    Alias ="LiquorRetail"
    Expression ="IIf(Item.SalesDepartmentID=1,Involine.price,0)"
    Alias ="LiquorCost"
    Expression ="IIf(Item.SalesDepartmentID=1,Involine.Cost,0)"
    Alias ="NonLiquorQuantity"
    Expression ="IIf(Item.SalesDepartmentID<>1,[Quantity],0)"
    Alias ="NonLiquorRetail"
    Expression ="IIf(Item.SalesDepartmentID<>1,Involine.price,0)"
    Alias ="NonLiquorCost"
    Expression ="IIf(Item.SalesDepartmentID<>1,Involine.Cost,0)"
    Alias ="InOrOut"
    Expression ="IIf([LiquorRetail]<0 Or [NonLiquorRetail]<0,\"Out\",\"In\")"
    Alias ="TotalQuanity"
    Expression ="InvoLine.Quantity"
    Alias ="TotalRetail"
    Expression ="InvoLine.Price"
    Alias ="TotalCost"
    Expression ="InvoLine.Cost"
    Expression ="Item.ID"
    Expression ="Item.Description"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
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
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorRetail"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorRetail"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalRetail"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiquorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NonLiquorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InOrOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQuanity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
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
    Bottom =235
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =195
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =318
        Top =4
        Right =414
        Bottom =118
        Top =0
        Name ="Item"
        Name =""
    End
End
