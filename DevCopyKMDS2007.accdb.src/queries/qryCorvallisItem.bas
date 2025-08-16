Operation =3
Name ="item"
Option =0
Where ="(((Trim([ItemCorv1].[description])) Is Not Null))"
Begin InputTables
    Name ="ItemCorv1"
    Name ="ItemCorv2"
End
Begin OutputColumns
    Alias ="ID"
    Name ="ID"
    Expression ="CLng(ItemCorv1.sku)"
    Alias ="descrip"
    Name ="Description"
    Expression ="Trim(ItemCorv1.description)"
    Alias ="StockNumber"
    Name ="StockNumber"
    Expression ="Trim([mfg part])"
    Alias ="SalesDeptID"
    Name ="SalesDepartmentID"
    Expression ="CInt([dept])"
    Alias ="Expr1"
    Name ="Location"
    Expression ="ItemCorv1.location"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="CLng([vendor number])"
    Alias ="Expr2"
    Name ="Onhand"
    Expression ="ItemCorv1.OnHand"
    Alias ="Expr3"
    Name ="OrderAt"
    Expression ="ItemCorv1.OrderAt"
    Alias ="Expr4"
    Name ="OrderTo"
    Expression ="ItemCorv1.OrderTo"
    Alias ="Expr5"
    Name ="Price"
    Expression ="ItemCorv1.Price"
    Alias ="Expr6"
    Name ="Cost"
    Expression ="ItemCorv1.Cost"
    Alias ="UnitsPerCase"
    Name ="UnitsPerCase"
    Expression ="ItemCorv1.[case q]"
    Alias ="UnitsPerOrder"
    Name ="UnitsPerOrder"
    Expression ="ItemCorv1.[case q]"
    Alias ="Expr7"
    Name ="DateLastSold"
    Expression ="ItemCorv2.DateLastSold"
    Alias ="dt"
    Name ="SetupDate"
    Expression ="4/1/2005"
    Alias ="dt2"
    Name ="ChangeDate"
    Expression ="4/1/2005"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemCorv1.description"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemCorv1.location"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="descrip"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StockNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDeptID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
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
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dt2"
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
    Bottom =229
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =151
        Bottom =45
        Top =0
        Name ="ItemCorv1"
        Name =""
    End
    Begin
        Left =254
        Top =14
        Right =390
        Bottom =53
        Top =0
        Name ="ItemCorv2"
        Name =""
    End
End
