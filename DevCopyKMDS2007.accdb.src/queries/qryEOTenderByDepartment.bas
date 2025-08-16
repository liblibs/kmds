Operation =1
Option =0
Where ="(((InvoLine5.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Tender"
End
Begin OutputColumns
    Alias ="Department"
    Expression ="Tender.ID"
    Alias ="DepartmentDesc"
    Expression ="Tender.Description"
    Alias ="TotalQ"
    Expression ="Sum(IIf([Quantity]=1,1,0))"
    Alias ="TotalP"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf([CashsysID]=1 And [Quantity]=1,1,0))"
    Alias ="D1"
    Expression ="Sum(IIf([CashsysID]=1,InvoLine5.Price,0))"
    Alias ="Q2"
    Expression ="Sum(IIf([CashsysID]=2 And [Quantity]=1,1,0))"
    Alias ="D2"
    Expression ="Sum(IIf([CashsysID]=2,InvoLine5.Price,0))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3 And [Quantity]=1,1,0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,InvoLine5.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4 And [Quantity]=1,1,0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,InvoLine5.Price,0))"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Tender.ID"
    Flag =0
End
Begin Groups
    Expression ="Tender.ID"
    GroupLevel =0
    Expression ="Tender.Description"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
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
    Bottom =245
    Left =0
    Top =9
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =195
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =251
        Top =66
        Right =347
        Bottom =180
        Top =0
        Name ="Tender"
        Name =""
    End
End
