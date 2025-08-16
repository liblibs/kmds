Operation =1
Option =0
Where ="((([Items].[Description]) Like 'M*'))"
Begin InputTables
    Name ="Items"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Items.Code"
    Alias ="Expr2"
    Expression ="Items.Description"
    Alias ="Expr3"
    Expression ="Items.Onhand"
    Alias ="Expr4"
    Expression ="Items.Price"
    Alias ="Expr5"
    Expression ="Items.PriceGroupID"
    Alias ="Expr6"
    Expression ="Items.SalesDepartmentID"
    Alias ="Expr7"
    Expression ="Items.StartDate"
    Alias ="Expr8"
    Expression ="Items.EndDate"
    Alias ="Expr9"
    Expression ="Items.NewPrice"
End
Begin OrderBy
    Expression ="Items.Description"
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
        dbText "Name" ="Expr1"
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
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
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
        Bottom =45
        Top =0
        Name ="Items"
        Name =""
    End
End
