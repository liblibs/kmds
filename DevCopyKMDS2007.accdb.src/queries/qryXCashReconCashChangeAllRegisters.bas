Operation =1
Option =0
Where ="(((InvoLineAllRegisters.TransType)='T' Or (InvoLineAllRegisters.TransType)='T'))"
Having ="(((InvoLineAllRegisters.Code)=1) AND ((InvoLineAllRegisters.InputCode)=\"Cash\")"
    ") OR (((InvoLineAllRegisters.Code)=1) AND ((InvoLineAllRegisters.InputCode)=\"Ch"
    "ange\"))"
Begin InputTables
    Name ="InvoLineAllRegisters"
    Name ="Tender"
End
Begin OutputColumns
    Expression ="InvoLineAllRegisters.Code"
    Alias ="Transact"
    Expression ="InvoLineAllRegisters.InputCode"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLineAllRegisters.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLineAllRegisters.Price)"
End
Begin Joins
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Tender"
    Expression ="InvoLineAllRegisters.Code = Tender.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLineAllRegisters.Code"
    Flag =0
End
Begin Groups
    Expression ="InvoLineAllRegisters.Code"
    GroupLevel =0
    Expression ="InvoLineAllRegisters.InputCode"
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
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="InvoLineAllRegisters"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =90
        Top =0
        Name ="Tender"
        Name =""
    End
End
