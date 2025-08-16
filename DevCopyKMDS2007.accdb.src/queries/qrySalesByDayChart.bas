Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((InvoLine.Trans"
    "Type)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\") And ((I"
    "nvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms"
    "!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="SalesDepartment"
    Name ="SalesGroup"
End
Begin OutputColumns
    Alias ="Week"
    Expression ="Format(InvoLine.SetupDate,\"ww\")"
    Alias ="Date"
    Expression ="Min(Format(InvoLine.SetupDate,\"mm/dd/yyyy\"))"
    Alias ="SumOfPrice"
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
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="SalesGroup"
    Expression ="SalesDepartment.SalesGroupID = SalesGroup.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
Begin Groups
    Expression ="Format(InvoLine.SetupDate,\"ww\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week"
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =201
        Top =40
        Right =297
        Bottom =154
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =389
        Top =41
        Right =485
        Bottom =155
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =571
        Top =13
        Right =667
        Bottom =127
        Top =0
        Name ="SalesGroup"
        Name =""
    End
End
