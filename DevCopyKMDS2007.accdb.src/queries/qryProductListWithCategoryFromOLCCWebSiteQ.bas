Operation =2
Name ="qryProductListWithCategoriesFromOLCCWebsite"
Option =8
Begin InputTables
    Name ="ProductListWithCategoriesFromOLCCWebsite"
End
Begin OutputColumns
    Alias ="Code"
    Expression ="CLng(Left(Trim([f1]),Len(Trim([f1]))-1) & l2n(Right(Trim([F1]),1)))"
    Alias ="Description"
    Expression ="Right([f2],(Len([f2])-InStr([f2],\" \")))"
    Alias ="Price"
    Expression ="CCur([f3])"
    Alias ="CaseQ"
    Expression ="CCur(Left([f2],InStr([f2],\" \")))/CCur([f3])"
    Alias ="Size"
    Expression ="L2N(Right([F1],1))"
    Alias ="Expr1"
    Expression ="ProductListWithCategoriesFromOLCCWebsite.F1"
    Alias ="ln"
    Expression ="l2n(Right(Trim([F1]),1))"
    Alias ="Expr2"
    Expression ="ProductListWithCategoriesFromOLCCWebsite.F5"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="6015"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CaseQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ln"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
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
        Right =407
        Bottom =120
        Top =0
        Name ="ProductListWithCategoriesFromOLCCWebsite"
        Name =""
    End
End
