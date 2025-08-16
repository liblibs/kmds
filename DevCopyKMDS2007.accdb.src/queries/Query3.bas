Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"C\"))"
Having ="(((Format([SetupDate],\"yyyymm\"))=Format(Now(),\"yyyymm\")))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="YearMonth"
    Expression ="Format([SetupDate],\"yyyymm\")"
    Alias ="ClubId"
    Expression ="InvoLine.SalesDept"
    Alias ="Weex"
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    Alias ="Week1"
    Expression ="Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now(),\"yyyy\")),\"ww\")"
    Expression ="InvoLine.Code"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="Week"
    Expression ="Sum(IIf(Format([SetupDate],\"w\")=1,[quantity],0))"
    Expression ="InvoLine.SetupDate"
End
Begin OrderBy
    Expression ="InvoLine.SalesDept"
    Flag =0
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    Flag =0
End
Begin Groups
    Expression ="Format([SetupDate],\"yyyymm\")"
    GroupLevel =0
    Expression ="InvoLine.SalesDept"
    GroupLevel =0
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    GroupLevel =0
    Expression ="Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now(),\"yyyy\")),\"ww\")"
    GroupLevel =0
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="InvoLine.SetupDate"
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="4860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YearMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Weex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week"
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
    Bottom =236
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =251
        Bottom =218
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
