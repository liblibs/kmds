Operation =2
Name ="OLCCProd"
Option =0
Begin InputTables
    Name ="ProductListFromOLCCWebsite"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="ProductListFromOLCCWebsite.F1"
    Alias ="Expr2"
    Expression ="ProductListFromOLCCWebsite.F3"
    Alias ="Liq"
    Expression ="CLng(Left([f1],Len([f1])-1) & l2n(Right([f1],1)))"
    Alias ="Size"
    Expression ="Right([f1],1)"
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Liq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
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
        Bottom =105
        Top =0
        Name ="ProductListFromOLCCWebsite"
        Name =""
    End
End
