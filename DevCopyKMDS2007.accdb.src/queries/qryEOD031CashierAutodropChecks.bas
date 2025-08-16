Operation =3
Name ="InvoLine5"
Option =0
Where ="(((Sysfile.AutoDropChecks)=True) AND ((InvoLine5.InputCode)<>\"Drop AutoDrop\"))"
Having ="(((Sum(InvoLine5.Price))<>0) AND ((InvoLine5.TransType)=\"T\") AND ((InvoLine5.C"
    "ode)=2))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Sysfile"
End
Begin OutputColumns
    Name ="InvoHeadID"
    Expression ="InvoLine5.Code"
    Name ="ID"
    Expression ="InvoLine5.CashierID"
    Name ="CashSysID"
    Expression ="InvoLine5.CashSysID"
    Name ="CashierID"
    Expression ="InvoLine5.CashierID"
    Name ="Code"
    Expression ="InvoLine5.Code"
    Alias ="SumOfPrice"
    Name ="Price"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="LastOfSetupDate"
    Name ="SetupDate"
    Expression ="Last(InvoLine5.SetupDate)"
    Name ="TransType"
    Expression ="InvoLine5.TransType"
    Alias ="InputCod"
    Name ="InputCode"
    Expression ="\"Drop AutoDrop\""
    Alias ="Nquantity"
    Name ="Quantity"
    Expression ="1"
End
Begin Groups
    Expression ="InvoLine5.CashSysID"
    GroupLevel =0
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="InvoLine5.Code"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
    GroupLevel =0
    Expression ="\"Drop AutoDrop\""
    GroupLevel =0
    Expression ="1"
    GroupLevel =0
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="InvoLine5.Code"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="LastOfSetupDate"
        dbInteger "ColumnWidth" ="3915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCod"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Nquantity"
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
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =191
        Bottom =120
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =229
        Top =6
        Right =325
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
