Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1 Or (Item.SalesDepartmentID)=2) AND ((InvoLine5.Tran"
    "sType)<>\"T\"))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Counter"
    Expression ="Sum(IIf([transType]=\"S\",InvoLine5.Price,0))"
    Alias ="ClubsGross"
    Expression ="Sum(IIf([transType]=\"C\",InvoLine5.Price,0))"
    Alias ="ClubsDiscount"
    Expression ="Sum(IIf([transType]=\"D\",InvoLine5.Price,0))"
    Alias ="ClubsNet"
    Expression ="Sum(IIf([transType]<>\"S\",InvoLine5.Price,0))"
    Alias ="OLCCDeposit"
    Expression ="Sum(InvoLine5.Price)"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
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
        dbText "Name" ="Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubsGross"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubsDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubsNet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCDeposit"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =188
        Top =5
        Right =284
        Bottom =119
        Top =0
        Name ="Item"
        Name =""
    End
End
