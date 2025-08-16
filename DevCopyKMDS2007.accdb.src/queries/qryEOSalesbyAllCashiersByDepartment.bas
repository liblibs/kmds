Operation =1
Option =0
Where ="(((InvoLine5.InputCode)<>\"Cash\" And (InvoLine5.InputCode)<>\"Change\" And (Inv"
    "oLine5.InputCode)<>\"Checks\" And (InvoLine5.InputCode)<>\"Visa MC\" And (InvoLi"
    "ne5.InputCode) Not Like \"Auth:*\") AND ((InvoLine5.Price)<>0))"
Having ="(((InvoLine5.TransType)<>\"X\"))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="InvoLine5"
    Name ="Item"
    Name ="Cashier"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="SalesGroup.GroupLevel"
    Expression ="InvoLine5.CashierID"
    Expression ="Cashier.Name"
    Alias ="SalesOrTender"
    Expression ="IIf([TransType]=\"T\",\"Loans and Drops\",\" Sales\")"
    Alias ="Group"
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.ID)"
    Alias ="GroupDesc"
    Expression ="IIf(IsNull(SalesGroup.Description),\"Loans and Drops\",SalesGroup.Description)"
    Alias ="Department"
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    Alias ="DepartmentDesc"
    Expression ="IIf(IsNull(SalesDepartment.Description),Tender.Description,IIf(Involine5.Transty"
        "pe='C',SalesDepartment.Description & \" Club Sales\",SalesDepartment.Description"
        "))"
    Alias ="TotalQ"
    Expression ="Sum(IIf([TransType]=\"T\",0,[Quantity]))"
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
    Alias ="Q5"
    Expression ="Sum(IIf([CashsysID]=5,[Quantity],0))"
    Alias ="D5"
    Expression ="Sum(IIf([CashsysID]=5,[InvoLine5].[Price],0))"
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
    RightTable ="Cashier"
    Expression ="InvoLine5.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Tender"
    Expression ="InvoLine5.Code = Tender.ID"
    Flag =2
End
Begin OrderBy
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.ID)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    Flag =0
End
Begin Groups
    Expression ="SalesGroup.GroupLevel"
    GroupLevel =0
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="IIf([TransType]=\"T\",\"Loans and Drops\",\" Sales\")"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.ID)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.Description),\"Loans and Drops\",SalesGroup.Description)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.Description),Tender.Description,IIf(Involine5.Transty"
        "pe='C',SalesDepartment.Description & \" Club Sales\",SalesDepartment.Description"
        "))"
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
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesOrTender"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesGroup.GroupLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D5"
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
    Bottom =228
    Left =0
    Top =16
    ColumnsShown =543
    Begin
        Left =638
        Top =31
        Right =734
        Bottom =145
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =476
        Top =16
        Right =572
        Bottom =130
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =200
        Top =7
        Right =296
        Bottom =226
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =342
        Top =14
        Right =438
        Bottom =128
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =32
        Top =52
        Right =128
        Bottom =166
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =339
        Top =140
        Right =435
        Bottom =254
        Top =0
        Name ="Tender"
        Name =""
    End
End
