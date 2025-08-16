Operation =1
Option =0
Where ="(((Item.SalesDepartmentID) Between Forms!PrintInventoryReport!cboSalesDepartment"
    "IDBegin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd Or (Item.SalesDep"
    "artmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Form"
    "s!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((InvoLine.Drawer)=0))"
Having ="(((First(InvoLine.SetupDate)) Between Forms!PrintInventoryReport!txtDateRangeFro"
    "m And Forms!PrintInventoryReport!txtDateRangeTo) And ((InvoLine.TransType)=\"X\""
    "))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Alias ="FirstOfSetupDate"
    Expression ="First(InvoLine.SetupDate)"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.ShiftID"
    Expression ="InvoLine.InputCode"
    Alias ="LiquorQuantity"
    Expression ="Sum(IIf(Item.SalesDepartmentID=1,[Quantity],0))"
    Alias ="LiquorRetail"
    Expression ="Sum(IIf(Item.SalesDepartmentID=1,Involine.price,0))"
    Alias ="LiquorCost"
    Expression ="Sum(IIf(Item.SalesDepartmentID=1,Involine.Cost,0))"
    Alias ="NonLiquorQuantity"
    Expression ="Sum(IIf(Item.SalesDepartmentID<>1,[Quantity],0))"
    Alias ="NonLiquorRetail"
    Expression ="Sum(IIf(Item.SalesDepartmentID<>1,Involine.price,0))"
    Alias ="NonLiquorCost"
    Expression ="Sum(IIf(Item.SalesDepartmentID<>1,Involine.Cost,0))"
    Alias ="InOrOut"
    Expression ="IIf([LiquorRetail]<0 Or [NonLiquorRetail]<0,\"Out\",\"In\")"
    Alias ="TotalQuanity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalRetail"
    Expression ="Sum(InvoLine.Price)"
    Alias ="TotalCost"
    Expression ="Sum(InvoLine.Cost)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="First(InvoLine.SetupDate)"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="InvoLine.ShiftID"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
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
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfSetupDate"
        dbInteger "ColumnWidth" ="2100"
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
    ColumnsShown =543
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
