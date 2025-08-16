Operation =3
Name ="InvoLine5"
Option =0
Where ="(((InvoLine.SetupDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]"
    " And [Forms]![PrintInventoryReport]![txtDateRangeTo]))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Name ="SetupDate"
    Expression ="InvoLine.SetupDate"
    Name ="InvoHeadID"
    Expression ="InvoLine.InvoHeadID"
    Name ="ID"
    Expression ="InvoLine.ID"
    Name ="CashSysID"
    Expression ="InvoLine.CashSysID"
    Name ="CashierID"
    Expression ="InvoLine.CashierID"
    Name ="ShiftID"
    Expression ="InvoLine.ShiftID"
    Name ="Code"
    Expression ="InvoLine.Code"
    Name ="SalesDept"
    Expression ="InvoLine.SalesDept"
    Name ="Quantity"
    Expression ="InvoLine.Quantity"
    Name ="Price"
    Expression ="InvoLine.Price"
    Name ="InputCode"
    Expression ="InvoLine.InputCode"
    Name ="SalesGroup"
    Expression ="InvoLine.SalesGroup"
    Name ="Drawer"
    Expression ="InvoLine.Drawer"
    Name ="TransType"
    Expression ="InvoLine.TransType"
    Name ="UpdateOnhand"
    Expression ="InvoLine.UpdateOnhand"
    Name ="Cost"
    Expression ="InvoLine.Cost"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
