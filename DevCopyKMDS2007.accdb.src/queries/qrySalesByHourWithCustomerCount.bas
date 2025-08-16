Operation =1
Option =0
Begin InputTables
    Name ="qrySalesByHour"
    Name ="qrySalesByHourInvoiceTotals"
End
Begin OutputColumns
    Expression ="qrySalesByHour.Date"
    Expression ="qrySalesByHour.SumOfPrice"
    Expression ="qrySalesByHour.[10 AM]"
    Expression ="qrySalesByHour.[11 AM]"
    Expression ="qrySalesByHour.[12 AM]"
    Expression ="qrySalesByHour.[1 PM]"
    Expression ="qrySalesByHour.[2 PM]"
    Expression ="qrySalesByHour.[3 PM]"
    Expression ="qrySalesByHour.[4 PM]"
    Expression ="qrySalesByHour.[5 PM]"
    Expression ="qrySalesByHour.[6 PM]"
    Expression ="qrySalesByHour.[7 PM]"
    Expression ="qrySalesByHour.[8 PM]"
    Expression ="qrySalesByHour.[9 PM]"
    Expression ="qrySalesByHour.Licensee"
    Alias ="LineCount"
    Expression ="qrySalesByHour.Count"
    Alias ="Count"
    Expression ="qrySalesByHourInvoiceTotals.CustomerCount"
End
Begin Joins
    LeftTable ="qrySalesByHour"
    RightTable ="qrySalesByHourInvoiceTotals"
    Expression ="qrySalesByHour.Date = qrySalesByHourInvoiceTotals.BusinessDate"
    Flag =2
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
        dbText "Name" ="LineCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
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
        Right =134
        Bottom =120
        Top =0
        Name ="qrySalesByHour"
        Name =""
    End
    Begin
        Left =177
        Top =-12
        Right =287
        Bottom =102
        Top =0
        Name ="qrySalesByHourInvoiceTotals"
        Name =""
    End
End
