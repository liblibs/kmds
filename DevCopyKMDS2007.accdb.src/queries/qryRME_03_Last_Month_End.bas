Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND (([Onhand]+IIf(IsNull([delta]),0,[delta]))>0))"
Begin InputTables
    Name ="Item"
    Name ="ResendMonthEnd"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.ID"
    Expression ="Item.Onhand"
    Expression ="ResendMonthEnd.delta"
    Alias ="MonthEndOnhand"
    Expression ="[Onhand]+IIf(IsNull([delta]),0,[delta])"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="ResendMonthEnd"
    Expression ="Item.ID = ResendMonthEnd.Code"
    Flag =2
End
Begin OrderBy
    Expression ="Item.ID"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthEndOnhand"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ResendMonthEnd.delta"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =60
    Right =1214
    Bottom =753
    Left =-1
    Top =-1
    Right =1138
    Bottom =320
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =251
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =522
        Top =12
        Right =666
        Bottom =312
        Top =0
        Name ="ResendMonthEnd"
        Name =""
    End
End
