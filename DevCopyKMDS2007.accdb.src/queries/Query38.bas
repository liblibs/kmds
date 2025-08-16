Operation =1
Option =0
Begin InputTables
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="SubCategory.ID"
    Expression ="SubCategory.CategoryID"
    Expression ="SubCategory.Description"
    Expression ="SubCategory.TargetGP"
    Expression ="SubCategory.Setup"
    Expression ="SubCategory.Change"
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
        dbText "Name" ="SubCategory.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.CategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.TargetGP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Setup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategory.Change"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1487
    Bottom =523
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
