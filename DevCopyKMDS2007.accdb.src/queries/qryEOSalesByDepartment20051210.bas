Operation =1
Option =0
Where ="(((InvoLine5.TransType)<>'X' And (InvoLine5.TransType)<>'T') AND ((InvoLine5.Inp"
    "utCode)<>\"Cash\" And (InvoLine5.InputCode)<>\"Change\" And (InvoLine5.InputCode"
    ")<>\"Checks\" And (InvoLine5.InputCode)<>\"Visa MC\") AND ((InvoLine5.Price)<>0)"
    ")"
Having ="(((IIf(IsNull([SalesDepartment].[ID]),[Tender].[Description],IIf([Involine5].[Tr"
    "anstype]='C',[SalesDepartment].[Description] & \" Club Sales\",[SalesDepartment]"
    ".[Description])))<>\"Visa MC\"))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="InvoLine5"
    Name ="Item"
    Name ="Tender"
End
Begin OutputColumns
    Alias ="GroupLvl"
    Expression ="IIf(IsNull(SalesGroup.id),\"2\",SalesGroup.GroupLevel)"
    Alias ="Group"
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Alias ="GroupDesc"
    Expression ="IIf(IsNull(SalesGroup.ID),\"Drops and Loans\",SalesGroup.Description)"
    Alias ="Department"
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    Alias ="DepartmentDesc"
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(Involine5.Transtype='C',Sa"
        "lesDepartment.Description & \" Club Sales\",SalesDepartment.Description))"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine5.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf([CashsysID]=1,[Quantity],0))"
    Alias ="D1"
    Expression ="Sum(IIf([CashsysID]=1,InvoLine5.Price,0))"
    Alias ="Q2"
    Expression ="Sum(IIf([CashsysID]=2,[Quantity],0))"
    Alias ="D2"
    Expression ="Sum(IIf([CashsysID]=2,InvoLine5.Price,0))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3,[Quantity],0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,InvoLine5.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4,[Quantity],0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,InvoLine5.Price,0))"
    Expression ="InvoLine5.TransType"
    Alias ="TotalC"
    Expression ="Sum(InvoLine5.Cost)"
    Alias ="TotalGP"
    Expression ="Sum(InvoLine5.Price-InvoLine5.Cost)"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =2
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =3
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =3
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =2
End
Begin OrderBy
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(Involine5.Transtype='C',Sa"
        "lesDepartment.Description & \" Club Sales\",SalesDepartment.Description))"
    Flag =0
End
Begin Groups
    Expression ="IIf(IsNull(SalesGroup.id),\"2\",SalesGroup.GroupLevel)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.ID),\"Drops and Loans\",SalesGroup.Description)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(Involine5.Transtype='C',Sa"
        "lesDepartment.Description & \" Club Sales\",SalesDepartment.Description))"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
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
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupLvl"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalGP"
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
    Top =16
    ColumnsShown =543
    Begin
        Left =449
        Top =16
        Right =545
        Bottom =130
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =307
        Top =29
        Right =403
        Bottom =143
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =28
        Top =-1
        Right =124
        Bottom =218
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =163
        Top =6
        Right =259
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =192
        Top =133
        Right =288
        Bottom =247
        Top =0
        Name ="Tender"
        Name =""
    End
End
