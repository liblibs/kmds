Operation =3
Name ="ItemMovement"
Option =0
Where ="(((ItemMovement.Description) Like \" Sales*\"))"
Begin InputTables
    Name ="ItemMovement"
End
Begin OutputColumns
    Name ="Code"
    Expression ="ItemMovement.Code"
    Alias ="Descr"
    Name ="Description"
    Expression ="\"Sales - Total\""
    Alias ="SumOfCurrMM"
    Name ="CurrMM"
    Expression ="Sum(ItemMovement.CurrMM)"
    Alias ="SumOfYTDQty"
    Name ="YTDQty"
    Expression ="Sum(ItemMovement.YTDQty)"
    Alias ="SumOfJan"
    Name ="Jan"
    Expression ="Sum(ItemMovement.Jan)"
    Alias ="SumOfFeb"
    Name ="Feb"
    Expression ="Sum(ItemMovement.Feb)"
    Alias ="SumOfMar"
    Name ="Mar"
    Expression ="Sum(ItemMovement.Mar)"
    Alias ="SumOfApr"
    Name ="Apr"
    Expression ="Sum(ItemMovement.Apr)"
    Alias ="SumOfMay"
    Name ="May"
    Expression ="Sum(ItemMovement.May)"
    Alias ="SumOfJun"
    Name ="Jun"
    Expression ="Sum(ItemMovement.Jun)"
    Alias ="SumOfJul"
    Name ="Jul"
    Expression ="Sum(ItemMovement.Jul)"
    Alias ="SumOfAug"
    Name ="Aug"
    Expression ="Sum(ItemMovement.Aug)"
    Alias ="SumOfSep"
    Name ="Sep"
    Expression ="Sum(ItemMovement.Sep)"
    Alias ="SumOfOct"
    Name ="Oct"
    Expression ="Sum(ItemMovement.Oct)"
    Alias ="SumOfNov"
    Name ="Nov"
    Expression ="Sum(ItemMovement.Nov)"
    Alias ="SumOfDec"
    Name ="Dec"
    Expression ="Sum(ItemMovement.Dec)"
End
Begin Groups
    Expression ="ItemMovement.Code"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemMovement.Description"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Descr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCurrMM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfYTDQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfJan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfFeb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfMar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfApr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfMay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfJun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfJul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfAug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfOct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfNov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfDec"
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
    Bottom =255
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =286
        Bottom =240
        Top =0
        Name ="ItemMovement"
        Name =""
    End
End
