Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="qryADS_ItemMovementWeightedByVendorADS"
End
Begin OutputColumns
    Name ="Item.AverageDailySales"
    Expression ="qryADS_ItemMovementWeightedByVendorADS.WADSCurrMM+qryADS_ItemMovementWeightedByV"
        "endorADS.WADSJan+qryADS_ItemMovementWeightedByVendorADS.WADSFeb+qryADS_ItemMovem"
        "entWeightedByVendorADS.WADSMar+qryADS_ItemMovementWeightedByVendorADS.WADSApr+qr"
        "yADS_ItemMovementWeightedByVendorADS.WADSMay+qryADS_ItemMovementWeightedByVendor"
        "ADS.WADSJun+qryADS_ItemMovementWeightedByVendorADS.WADSJul+qryADS_ItemMovementWe"
        "ightedByVendorADS.WADSAug+qryADS_ItemMovementWeightedByVendorADS.WADSSep+qryADS_"
        "ItemMovementWeightedByVendorADS.WADSOct+qryADS_ItemMovementWeightedByVendorADS.W"
        "ADSNov+qryADS_ItemMovementWeightedByVendorADS.WADSDec"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryADS_ItemMovementWeightedByVendorADS"
    Expression ="Item.ID = qryADS_ItemMovementWeightedByVendorADS.Code"
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
    Bottom =93
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
        Left =172
        Top =6
        Right =559
        Bottom =120
        Top =0
        Name ="qryADS_ItemMovementWeightedByVendorADS"
        Name =""
    End
End
