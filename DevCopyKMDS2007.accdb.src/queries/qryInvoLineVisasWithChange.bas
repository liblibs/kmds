Operation =1
Option =0
Where ="(((InvoLine.Price)<>0) AND ((InvoLine.InputCode)=\"Change\") AND ((InvoLine.Tran"
    "sType)=\"T\") AND ((InvoLine.Code)<>3))"
Begin InputTables
    Name ="qryInvoLineVisas"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.InvoHeadID"
    Expression ="qryInvoLineVisas.Price"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InputCode"
    Expression ="InvoLine.CashierID"
End
Begin Joins
    LeftTable ="qryInvoLineVisas"
    RightTable ="InvoLine"
    Expression ="qryInvoLineVisas.CashSysID = InvoLine.CashSysID"
    Flag =1
    LeftTable ="qryInvoLineVisas"
    RightTable ="InvoLine"
    Expression ="qryInvoLineVisas.InvoHeadID = InvoLine.InvoHeadID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =0
    Expression ="InvoLine.CashSysID"
    Flag =0
    Expression ="InvoLine.InvoHeadID"
    Flag =0
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInvoLineVisas.Price"
        dbInteger "ColumnWidth" ="1935"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =179
        Bottom =120
        Top =0
        Name ="qryInvoLineVisas"
        Name =""
    End
    Begin
        Left =305
        Top =2
        Right =528
        Bottom =116
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
