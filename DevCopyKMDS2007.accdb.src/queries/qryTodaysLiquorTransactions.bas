Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND ((InvoLine.SetupDate)>Date()))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Today"
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
    Alias ="Purchases"
    Expression ="Sum(IIf([transType]=\"P\",InvoLine.Price,0))"
    Alias ="TransfersOUT"
    Expression ="Sum(IIf([transType]=\"X\" And InvoLine.Price<0,InvoLine.Price,0))"
    Alias ="TransfersIN"
    Expression ="Sum(IIf([transType]=\"X\" And InvoLine.Price>0,InvoLine.Price,0))"
    Alias ="Adjustments"
    Expression ="Sum(IIf([transType]=\"A\",InvoLine.Price,0))"
    Alias ="Counter"
    Expression ="Sum(IIf([transType]=\"S\",0-InvoLine.Price,0))"
    Alias ="Clubs"
    Expression ="Sum(IIf([transType]=\"C\",0-InvoLine.Price,0))"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
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
        dbText "Name" ="Purchases"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Today"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Adjustments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransfersOUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransfersIN"
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
    Bottom =322
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =280
        Bottom =285
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =321
        Top =65
        Right =417
        Bottom =179
        Top =0
        Name ="Item"
        Name =""
    End
End
