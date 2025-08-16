Operation =1
Option =0
Where ="(((InvoLineAllRegisters.TransType)<>'X') AND ((InvoLineAllRegisters.InputCode)<>"
    "\"Cash\" And (InvoLineAllRegisters.InputCode)<>\"Change\" And (InvoLineAllRegist"
    "ers.InputCode)<>\"Checks\" And (InvoLineAllRegisters.InputCode)<>\"Visa MC\" And"
    " (InvoLineAllRegisters.InputCode) Not Like \"Auth:*\") AND ((InvoLineAllRegister"
    "s.Price)<>0))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="InvoLineAllRegisters"
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
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(InvoLineAllRegisters.Trans"
        "type='C',SalesDepartment.Description & \" Club Sales\",SalesDepartment.Descripti"
        "on))"
    Alias ="TotalQ"
    Expression ="Sum(IIf([TransType]=\"T\",1,[Quantity]))"
    Alias ="TotalP"
    Expression ="Sum(InvoLineAllRegisters.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf([CashsysID]=1,IIf([TransType]=\"T\",1,[Quantity]),0))"
    Alias ="D1"
    Expression ="Sum(IIf([CashsysID]=1,InvoLineAllRegisters.Price,0))"
    Alias ="Q2"
    Expression ="Sum(IIf([CashsysID]=2,IIf([TransType]=\"T\",1,[Quantity]),0))"
    Alias ="D2"
    Expression ="Sum(IIf([CashsysID]=2,InvoLineAllRegisters.Price,0))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3,IIf([TransType]=\"T\",1,[Quantity]),0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,InvoLineAllRegisters.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4,IIf([TransType]=\"T\",1,[Quantity]),0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,InvoLineAllRegisters.Price,0))"
    Expression ="InvoLineAllRegisters.TransType"
    Alias ="TotalC"
    Expression ="0"
    Alias ="TotalGP"
    Expression ="Sum(InvoLineAllRegisters.Price)"
    Alias ="Q5"
    Expression ="Sum(IIf([CashsysID]=5,IIf([TransType]=\"T\",1,[Quantity]),0))"
    Alias ="D5"
    Expression ="Sum(IIf([CashsysID]=5,[InvoLineAllRegisters].[Price],0))"
End
Begin Joins
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Item"
    Expression ="InvoLineAllRegisters.Code = Item.ID"
    Flag =2
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =3
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =3
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Tender"
    Expression ="InvoLineAllRegisters.Code = Tender.ID"
    Flag =2
End
Begin OrderBy
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.ID,SalesDepartment.ID)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(InvoLineAllRegisters.Trans"
        "type='C',SalesDepartment.Description & \" Club Sales\",SalesDepartment.Descripti"
        "on))"
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
    Expression ="IIf(IsNull(SalesDepartment.ID),Tender.Description,IIf(InvoLineAllRegisters.Trans"
        "type='C',SalesDepartment.Description & \" Club Sales\",SalesDepartment.Descripti"
        "on))"
    GroupLevel =0
    Expression ="InvoLineAllRegisters.TransType"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbInteger "ColumnWidth" ="1620"
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
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="TotalC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalGP"
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
    Begin
        dbText "Name" ="InvoLineAllRegisters.TransType"
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLineAllRegisters"
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
