Operation =1
Option =0
Where ="(((SalesGroup.ID)=212) And ((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\""
    "C\" Or (InvoLine.TransType)=\"D\") And ((InvoLine.SetupDate) Between Forms!Print"
    "InventoryReport!txtDateRangeFrom And Forms!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="SalesGroup"
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Week"
    Expression ="(CLng(Format(InvoLine.SetupDate,\"yyyy\"))*100)+CLng(Format(InvoLine.SetupDate,\""
        "ww\"))"
    Alias ="BeginDate"
    Expression ="Min(Format(InvoLine.SetupDate,\"mm/dd/yyyy\"))"
    Alias ="TotalSales"
    Expression ="Sum(InvoLine.Price)"
    Alias ="Sun"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Sun\",InvoLine.PRICE,0))"
    Alias ="Mon"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Mon\",InvoLine.PRICE,0))"
    Alias ="Tue"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Tue\",InvoLine.PRICE,0))"
    Alias ="Wed"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Wed\",InvoLine.PRICE,0))"
    Alias ="Thu"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Thu\",InvoLine.PRICE,0))"
    Alias ="Fri"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Fri\",InvoLine.PRICE,0))"
    Alias ="Sat"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"ddd\")=\"Sat\",InvoLine.PRICE,0))"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="SalesGroup"
    Expression ="SalesDepartment.SalesGroupID = SalesGroup.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
End
Begin Groups
    Expression ="(CLng(Format(InvoLine.SetupDate,\"yyyy\"))*100)+CLng(Format(InvoLine.SetupDate,\""
        "ww\"))"
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
        dbText "Name" ="Week"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BeginDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mon"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Thu"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fri"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sat"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =435
        Top =116
        Right =531
        Bottom =230
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =641
        Top =9
        Right =737
        Bottom =123
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =140
        Bottom =270
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =186
        Top =5
        Right =353
        Bottom =254
        Top =0
        Name ="Item"
        Name =""
    End
End
