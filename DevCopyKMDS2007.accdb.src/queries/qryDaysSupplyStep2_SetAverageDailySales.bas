Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryDaysSupplyItemMovementWeightedADS"
End
Begin OutputColumns
    Name ="Item.AverageDailySales"
    Expression ="qryDaysSupplyItemMovementWeightedADS.WADSCurrMM+qryDaysSupplyItemMovementWeighte"
        "dADS.WADSJan+qryDaysSupplyItemMovementWeightedADS.WADSFeb+qryDaysSupplyItemMovem"
        "entWeightedADS.WADSMar+qryDaysSupplyItemMovementWeightedADS.WADSApr+qryDaysSuppl"
        "yItemMovementWeightedADS.WADSMay+qryDaysSupplyItemMovementWeightedADS.WADSJun+qr"
        "yDaysSupplyItemMovementWeightedADS.WADSJul+qryDaysSupplyItemMovementWeightedADS."
        "WADSAug+qryDaysSupplyItemMovementWeightedADS.WADSSep+qryDaysSupplyItemMovementWe"
        "ightedADS.WADSOct+qryDaysSupplyItemMovementWeightedADS.WADSNov+qryDaysSupplyItem"
        "MovementWeightedADS.WADSDec"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryDaysSupplyItemMovementWeightedADS"
    Expression ="Item.ID = qryDaysSupplyItemMovementWeightedADS.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Turns"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
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
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =691
        Bottom =120
        Top =0
        Name ="qryDaysSupplyItemMovementWeightedADS"
        Name =""
    End
End
