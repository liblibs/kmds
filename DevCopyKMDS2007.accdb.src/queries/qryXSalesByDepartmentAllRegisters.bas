Operation =1
Option =0
Where ="(((InvoLineAllRegisters.TransType)<>'T' And (InvoLineAllRegisters.TransType)<>'X"
    "'))"
Begin InputTables
    Name ="InvoLineAllRegisters"
    Name ="SalesDepartment"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="SalesDepartment.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLineAllRegisters.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLineAllRegisters.Price)"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =3
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Item"
    Expression ="InvoLineAllRegisters.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="Item.SalesDepartmentID"
    GroupLevel =0
    Expression ="SalesDepartment.Description"
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
    Bottom =245
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =93
        Top =50
        Right =189
        Bottom =257
        Top =0
        Name ="InvoLineAllRegisters"
        Name =""
    End
    Begin
        Left =543
        Top =52
        Right =639
        Bottom =166
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =301
        Top =56
        Right =397
        Bottom =170
        Top =0
        Name ="Item"
        Name =""
    End
End
