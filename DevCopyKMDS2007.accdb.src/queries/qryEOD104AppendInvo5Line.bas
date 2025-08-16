Operation =3
Name ="InvoLine"
Option =0
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Name ="InvoHeadID"
    Expression ="InvoLine5.InvoHeadID"
    Name ="ID"
    Expression ="InvoLine5.ID"
    Name ="CashSysID"
    Expression ="InvoLine5.CashSysID"
    Name ="CashierID"
    Expression ="InvoLine5.CashierID"
    Name ="ShiftID"
    Expression ="InvoLine5.ShiftID"
    Name ="Code"
    Expression ="InvoLine5.Code"
    Name ="SalesDept"
    Expression ="InvoLine5.SalesDept"
    Name ="Quantity"
    Expression ="InvoLine5.Quantity"
    Name ="Price"
    Expression ="InvoLine5.Price"
    Name ="InputCode"
    Expression ="InvoLine5.InputCode"
    Name ="SalesGroup"
    Expression ="InvoLine5.SalesGroup"
    Name ="Drawer"
    Expression ="InvoLine5.Drawer"
    Name ="TransType"
    Expression ="InvoLine5.TransType"
    Name ="UpdateOnhand"
    Expression ="InvoLine5.UpdateOnhand"
    Name ="SetupDate"
    Expression ="InvoLine5.SetupDate"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine5.InputCode"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine5"
        Name =""
    End
End
