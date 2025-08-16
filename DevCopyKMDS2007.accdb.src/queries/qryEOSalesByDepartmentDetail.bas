Operation =1
Option =0
Where ="(((InvoLine5.TransType)<>'X' And (InvoLine5.TransType)<>'T') AND ((InvoLine5.Inp"
    "utCode)<>\"Cash\" And (InvoLine5.InputCode)<>\"Change\" And (InvoLine5.InputCode"
    ")<>\"Checks\" And (InvoLine5.InputCode)<>\"Visa MC\") AND ((InvoLine5.Price)<>0)"
    ")"
Having ="(((Item.ID) Is Not Null))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Alias ="GroupLvl"
    Expression ="IIf(IsNull(SalesGroup.id),\"2\",SalesGroup.GroupLevel)"
    Alias ="Group"
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Alias ="GroupDesc"
    Expression ="IIf(IsNull(SalesGroup.ID),\"Drops and Loans\",SalesGroup.Description)"
    Alias ="Department"
    Expression ="IIf(IsNull(SalesDepartment.ID),0,SalesDepartment.ID)"
    Alias ="DepartmentDesc"
    Expression ="IIf(IsNull(SalesDepartment.ID),\"N/A\",IIf(Involine5.Transtype='C',SalesDepartme"
        "nt.Description & \" Club Sales\",SalesDepartment.Description))"
    Expression ="Item.Description"
    Expression ="Item.ID"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine5.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine5.Price)"
    Expression ="InvoLine5.TransType"
    Alias ="TotalC"
    Expression ="Sum(InvoLine5.Cost)"
    Alias ="TotalGP"
    Expression ="Sum(InvoLine5.Price-InvoLine5.Cost)"
    Alias ="TotalReturnsQ"
    Expression ="Sum(IIf([Quantity]<0,[Quantity],0))"
    Alias ="TotalReturnsD"
    Expression ="Sum(IIf([Quantity]<0,InvoLine5.price,0))"
    Alias ="TotalSalesQ"
    Expression ="Sum(IIf([Quantity]>0,[Quantity],0))"
    Alias ="TotalSalesD"
    Expression ="Sum(IIf([Quantity]>0,InvoLine5.price,0))"
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
End
Begin OrderBy
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),0,SalesDepartment.ID)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),\"N/A\",IIf(Involine5.Transtype='C',SalesDepartme"
        "nt.Description & \" Club Sales\",SalesDepartment.Description))"
    Flag =0
    Expression ="Item.ID"
    Flag =0
End
Begin Groups
    Expression ="IIf(IsNull(SalesGroup.id),\"2\",SalesGroup.GroupLevel)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesGroup.ID),\"Drops and Loans\",SalesGroup.Description)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.ID),0,SalesDepartment.ID)"
    GroupLevel =0
    Expression ="IIf(IsNull(SalesDepartment.ID),\"N/A\",IIf(Involine5.Transtype='C',SalesDepartme"
        "nt.Description & \" Club Sales\",SalesDepartment.Description))"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="Item.ID"
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
    Begin
        dbText "Name" ="TotalReturnsQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalReturnsD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSalesQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSalesD"
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
        Left =449
        Top =23
        Right =545
        Bottom =137
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =307
        Top =36
        Right =403
        Bottom =150
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =28
        Top =6
        Right =124
        Bottom =225
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =163
        Top =13
        Right =259
        Bottom =127
        Top =0
        Name ="Item"
        Name =""
    End
End
