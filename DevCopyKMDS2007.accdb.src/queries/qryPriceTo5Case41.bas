Operation =4
Option =0
Where ="(((Right(Format([newprice],\"Fixed\"),1))=4))"
Begin InputTables
    Name ="PricingWorkTable"
End
Begin OutputColumns
    Name ="PricingWorkTable.NewPrice"
    Expression ="[newprice]+0.01"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricingWorkTable.NewPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Right([newprice],1)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Right(Format([newprice],0),1)"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =179
    Top =36
    Right =913
    Bottom =842
    Left =-1
    Top =-1
    Right =702
    Bottom =403
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =272
        Top =0
        Name ="PricingWorkTable"
        Name =""
    End
End
