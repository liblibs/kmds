Operation =2
Name ="PalmTo"
Option =0
Begin InputTables
    Name ="Audit"
    Name ="Item"
    Name ="qryAuditFileToPalm_Step_2a"
End
Begin OutputColumns
    Alias ="Code"
    Expression ="Audit.ItemID"
    Expression ="Item.Description"
    Alias ="UPCCode"
    Expression ="IIf(IsNull([FirstOfID]),Audit.ItemID,[FirstOfID])"
    Expression ="Audit.OnShelf"
End
Begin Joins
    LeftTable ="Audit"
    RightTable ="Item"
    Expression ="Audit.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryAuditFileToPalm_Step_2a"
    Expression ="Item.ID = qryAuditFileToPalm_Step_2a.ItemID"
    Flag =2
End
Begin OrderBy
    Expression ="IIf(IsNull([FirstOfID]),Audit.ItemID,[FirstOfID])"
    Flag =0
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
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCCode"
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
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Audit"
        Name =""
    End
    Begin
        Left =192
        Top =21
        Right =288
        Bottom =135
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =460
        Top =6
        Right =635
        Bottom =105
        Top =0
        Name ="qryAuditFileToPalm_Step_2a"
        Name =""
    End
End
