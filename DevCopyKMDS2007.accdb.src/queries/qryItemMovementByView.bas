Operation =1
Option =0
Where ="(((Format(CDate([mm] & \"/01/\" & [yr]),\"yyyymm\"))>=Format((Now()-365),\"yyyym"
    "m\") And (Format(CDate([mm] & \"/01/\" & [yr]),\"yyyymm\"))<Format((Now()),\"yyy"
    "ymm\")))"
Begin InputTables
    Name ="dbo_vwInvoLineByCodeByYearMonth"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.Code"
    Alias ="YYYY"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.YR"
    Alias ="Expr2"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.MM"
    Alias ="Expr3"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.TransType"
    Alias ="SumOfQuantity"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.Qty"
    Alias ="SumOfPrice"
    Expression ="dbo_vwInvoLineByCodeByYearMonth.Sales"
    Alias ="TranMMYY"
    Expression ="Format(CDate([mm] & \"/01/\" & [yr]),\"yyyymm\")"
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
        dbText "Name" ="TranMMYY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YYYY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="dbo_vwInvoLineByCodeByYearMonth"
        Name =""
    End
End
