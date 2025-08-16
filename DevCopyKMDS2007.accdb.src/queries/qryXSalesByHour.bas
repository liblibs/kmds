Operation =1
Option =0
Where ="(((InvoLine.TransType)<>'T' And (InvoLine.TransType)<>'X') AND ((SalesDepartment"
    ".ID)<=50))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Alias ="DateHour"
    Expression ="Format([InvoLine]![SetupDate],\"mm-dd    hh:\"\"00 \"\"ampm\")"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Format([InvoLine]![SetupDate],\"mm-dd    hh:\"\"00 \"\"ampm\")"
    Flag =0
End
Begin Groups
    Expression ="Format([InvoLine]![SetupDate],\"mm-dd    hh:\"\"00 \"\"ampm\")"
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
        dbText "Name" ="DateHour"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =245
    Left =22
    Top =0
    ColumnsShown =543
    Begin
        Left =154
        Top =2
        Right =250
        Bottom =191
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =305
        Top =27
        Right =401
        Bottom =231
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =439
        Top =140
        Right =535
        Bottom =224
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
