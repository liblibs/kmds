Operation =1
Option =0
Where ="(((dbo_OLCCScanCode.UPCCode) Like \"10*\") AND ((Len([UPCCode]))=14))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
End
Begin OutputColumns
    Expression ="dbo_OLCCScanCode.ID"
    Expression ="dbo_OLCCScanCode.Code"
    Expression ="dbo_OLCCScanCode.UPCCode"
    Alias ="LID"
    Expression ="Len([UPCCode])"
    Alias ="ThirdDigit"
    Expression ="Mid([upccode],3,1)"
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
        dbText "Name" ="dbo_OLCCScanCode.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ThirdDigit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =63
    Top =160
    Right =1246
    Bottom =870
    Left =-1
    Top =-1
    Right =1160
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =74
        Top =44
        Right =218
        Bottom =188
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
End
