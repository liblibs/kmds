Operation =1
Option =0
Where ="(((InvoLine5.Code)=3) AND ((InvoLine5.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Alias ="Auth"
    Expression ="Right([InputCode],Len([InputCode])-5)"
    Expression ="InvoLine5.ID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.CashierID"
    Expression ="InvoLine5.ShiftID"
    Expression ="InvoLine5.Code"
    Expression ="InvoLine5.SalesDept"
    Expression ="InvoLine5.Quantity"
    Expression ="InvoLine5.Price"
    Expression ="InvoLine5.InputCode"
    Expression ="InvoLine5.SalesGroup"
    Expression ="InvoLine5.Drawer"
    Expression ="InvoLine5.TransType"
    Expression ="InvoLine5.UpdateOnhand"
    Expression ="InvoLine5.SetupDate"
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
        dbText "Name" ="Auth"
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
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =188
        Bottom =302
        Top =0
        Name ="InvoLine5"
        Name =""
    End
End
