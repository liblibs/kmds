Operation =1
Option =0
Where ="(((Format([SetupDate],\"yyyymm\"))=Format(Now(),\"yyyymm\")) AND ((InvoLine.Tran"
    "sType)=\"C\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="ClubId"
    Expression ="InvoLine.SalesDept"
    Alias ="Week"
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    Expression ="InvoLine.Code"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
End
Begin OrderBy
    Expression ="InvoLine.SalesDept"
    Flag =0
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.SalesDept"
    GroupLevel =0
    Expression ="(Format([SetupDate],\"ww\")-Format((Format(Now(),\"mm\") & \"/01/\" & Format(Now"
        "(),\"yyyy\")),\"ww\"))+1"
    GroupLevel =0
    Expression ="InvoLine.Code"
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
        dbText "Name" ="Week"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubId"
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
