Operation =1
Option =0
Begin InputTables
    Name ="MetrologicLookup"
End
Begin OutputColumns
    Expression ="MetrologicLookup.UPC"
    Expression ="MetrologicLookup.Descr"
    Expression ="MetrologicLookup.Item"
    Expression ="MetrologicLookup.PRI"
    Expression ="MetrologicLookup.curPrice"
    Expression ="MetrologicLookup.Shelf"
    Expression ="MetrologicLookup.Description"
    Alias ="UCase"
    Expression ="Format([UnitsPerCase],\"@@@\")"
End
Begin Parameters
    Name ="prmSalesdepartmentIdBegin"
    Flag =4
    Name ="prmSalesdepartmentIdEnd"
    Flag =4
    Name ="prmCategoryDescBegin"
    Flag =10
    Name ="prmCategoryDescEnd"
    Flag =10
End
Begin OrderBy
    Expression ="MetrologicLookup.UPC"
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
        dbText "Name" ="UCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.curPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.PRI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.Item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.Descr"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.UPC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicLookup.Shelf"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-363
    Top =95
    Right =1061
    Bottom =860
    Left =-1
    Top =-1
    Right =1392
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =331
        Top =2
        Right =475
        Bottom =146
        Top =0
        Name ="MetrologicLookup"
        Name =""
    End
End
