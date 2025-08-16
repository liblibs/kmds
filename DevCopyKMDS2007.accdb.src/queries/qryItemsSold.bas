Operation =1
Option =0
Where ="(((SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBe"
    "gin And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((Involine.Trans"
    "Type)=\"S\" Or (Involine.TransType)=\"C\") And ((Involine.Price)<>0) And ((Invol"
    "ine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom And Forms!Pri"
    "ntInventoryReport!txtDateRangeTo))"
Having ="(((Item.ID) Is Not Null))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Involine"
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
    Expression ="SalesDepartment.Description"
    Expression ="Item.Description"
    Expression ="Item.ID"
    Alias ="TotalQ"
    Expression ="Sum(Involine.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(Involine.Price)"
    Alias ="TotalC"
    Expression ="Sum(Involine.Cost)"
    Alias ="TotalGP"
    Expression ="Sum(Involine.Price-Involine.Cost)"
    Alias ="TotalReturnsQ"
    Expression ="Sum(IIf([Quantity]<0,[Quantity],0))"
    Alias ="TotalReturnsD"
    Expression ="Sum(IIf([Quantity]<0,Involine.price,0))"
    Alias ="TotalSalesQ"
    Expression ="Sum(IIf([Quantity]>0,[Quantity],0))"
    Alias ="TotalSalesD"
    Expression ="Sum(IIf([Quantity]>0,Involine.price,0))"
End
Begin Joins
    LeftTable ="Involine"
    RightTable ="Item"
    Expression ="Involine.Code = Item.ID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
End
Begin OrderBy
    Expression ="IIf(IsNull(SalesGroup.ID),999,SalesGroup.id)"
    Flag =0
    Expression ="IIf(IsNull(SalesDepartment.ID),0,SalesDepartment.ID)"
    Flag =0
    Expression ="SalesDepartment.Description"
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
    Expression ="SalesDepartment.Description"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="Item.ID"
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
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="2400"
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
        Left =565
        Top =18
        Right =661
        Bottom =132
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Involine"
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
