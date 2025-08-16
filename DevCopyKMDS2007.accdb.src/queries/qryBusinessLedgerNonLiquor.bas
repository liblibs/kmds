Operation =1
Option =0
Where ="(((SalesGroup.Description)<>\"Liquor\") And ((InvoLine.SetupDate) Between Forms!"
    "PrintInventoryReport!txtDateRangeFrom And Forms!PrintInventoryReport!txtDateRang"
    "eTo))"
Begin InputTables
    Name ="InvoLine"
    Name ="TranTypes"
    Name ="Item"
    Name ="SalesDepartment"
    Name ="SalesGroup"
End
Begin OutputColumns
    Alias ="BusDate"
    Expression ="Format(involine.SetupDate,\"mm/dd/yyyy\")"
    Alias ="SalesCounter"
    Expression ="Sum(IIf([Transtype]=\"S\",Involine.Price,0))"
    Alias ="SalesClubs"
    Expression ="Sum(IIf([Transtype]=\"C\",Involine.Price,0))"
    Alias ="Purchases"
    Expression ="Sum(IIf([Transtype]=\"P\",Involine.Price*Involine.Quantity,0))"
    Alias ="Transfers"
    Expression ="Sum(IIf([Transtype]=\"X\",Involine.Price,0))"
    Alias ="Adjust"
    Expression ="Sum(IIf([Transtype]=\"A\",Involine.Price*Involine.Quantity,0))"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType = TranTypes.Code"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="SalesGroup"
    Expression ="SalesDepartment.SalesGroupID = SalesGroup.ID"
    Flag =1
End
Begin Groups
    Expression ="Format(involine.SetupDate,\"mm/dd/yyyy\")"
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
        dbText "Name" ="BusDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesCounter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesClubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Purchases"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transfers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjust"
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
    Bottom =208
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =189
        Bottom =180
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =253
        Top =111
        Right =349
        Bottom =210
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =341
        Top =-5
        Right =437
        Bottom =109
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =495
        Top =6
        Right =591
        Bottom =120
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =629
        Top =6
        Right =725
        Bottom =120
        Top =0
        Name ="SalesGroup"
        Name =""
    End
End
