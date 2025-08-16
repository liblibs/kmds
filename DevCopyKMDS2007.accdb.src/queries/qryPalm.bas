Operation =2
Name ="aPalm"
Option =0
Where ="(((vwUPCItemsWithCurrentPricing.LenID)>7))"
Having ="(((vwUPCItemsWithCurrentPricing.ID) Between [BeginCode] And [EndCode]))"
Begin InputTables
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Alias ="FirstOfUPCCode"
    Expression ="First(vwUPCItemsWithCurrentPricing.UPCCode)"
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    Expression ="vwUPCItemsWithCurrentPricing.Description"
    Expression ="vwUPCItemsWithCurrentPricing.Onhand"
End
Begin OrderBy
    Expression ="First(vwUPCItemsWithCurrentPricing.UPCCode)"
    Flag =0
End
Begin Groups
    Expression ="vwUPCItemsWithCurrentPricing.ID"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.Description"
    GroupLevel =0
    Expression ="vwUPCItemsWithCurrentPricing.Onhand"
    GroupLevel =0
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
        dbText "Name" ="FirstOfUPCCode"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =228
        Bottom =120
        Top =0
        Name ="vwUPCItemsWithCurrentPricing"
        Name =""
    End
End
