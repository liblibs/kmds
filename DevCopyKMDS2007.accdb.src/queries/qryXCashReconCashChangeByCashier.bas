Operation =1
Option =0
Having ="(((InvoLine.Code)=1) AND ((InvoLine.InputCode)=\"Cash\") AND ((Cashier.ID)=[parC"
    "ashierNumber]) AND ((InvoLine.TransType)=\"T\")) OR (((InvoLine.Code)=1) AND ((I"
    "nvoLine.InputCode)=\"Change\") AND ((Cashier.ID)=[parCashierNumber]) AND ((InvoL"
    "ine.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoHead"
    Name ="InvoLine"
    Name ="Cashier"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="Transact"
    Expression ="InvoLine.InputCode"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="InvoHead"
    RightTable ="InvoLine"
    Expression ="InvoHead.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="InvoHead"
    RightTable ="InvoLine"
    Expression ="InvoHead.ID = InvoLine.InvoHeadID"
    Flag =1
    LeftTable ="InvoHead"
    RightTable ="Cashier"
    Expression ="InvoHead.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code = Tender.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="Cashier.ID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="InvoLine.TransType"
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
        dbText "Name" ="Transact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
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
        Left =24
        Top =0
        Right =120
        Bottom =99
        Top =0
        Name ="InvoHead"
        Name =""
    End
    Begin
        Left =290
        Top =3
        Right =386
        Bottom =102
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =173
        Top =71
        Right =269
        Bottom =155
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =473
        Top =19
        Right =569
        Bottom =103
        Top =0
        Name ="Tender"
        Name =""
    End
End
