Operation =6
Option =0
Begin InputTables
    Name ="qryItemMovement"
End
Begin OutputColumns
    Expression ="qryItemMovement.Code"
    GroupLevel =2
    Expression ="qryItemMovement.TransType"
    GroupLevel =2
    Expression ="qryItemMovement.MM"
    GroupLevel =1
    Alias ="SumOfSumOfQuantity"
    Expression ="Sum(qryItemMovement.SumOfQuantity)"
End
Begin Groups
    Expression ="qryItemMovement.Code"
    GroupLevel =2
    Expression ="qryItemMovement.TransType"
    GroupLevel =2
    Expression ="qryItemMovement.MM"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="200212"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200301"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200302"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200303"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200304"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200305"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200306"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200307"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200308"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200309"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200310"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200311"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="200312"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="01"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="02"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="03"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="04"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="05"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="06"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="07"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="08"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="09"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
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
    Bottom =212
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =38
        Top =6
        Right =287
        Bottom =150
        Top =0
        Name ="qryItemMovement"
        Name =""
    End
End
