Operation =1
Option =0
Having ="(((Min(PalmOrderGuideFrom.Physical)) Is Not Null And (Min(PalmOrderGuideFrom.Phy"
    "sical))>0))"
Begin InputTables
    Name ="PalmOrderGuideFrom"
    Name ="Item"
End
Begin OutputColumns
    Expression ="PalmOrderGuideFrom.ItemID"
    Alias ="MinOfPhysical"
    Expression ="Min(PalmOrderGuideFrom.Physical)"
    Alias ="Liquor"
    Expression ="IIf([SalesDepartmentID]=1,255,0)"
End
Begin Joins
    LeftTable ="PalmOrderGuideFrom"
    RightTable ="Item"
    Expression ="PalmOrderGuideFrom.ItemID = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="PalmOrderGuideFrom.ItemID"
    Flag =0
End
Begin Groups
    Expression ="PalmOrderGuideFrom.ItemID"
    GroupLevel =0
    Expression ="IIf([SalesDepartmentID]=1,255,0)"
    GroupLevel =0
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
        dbText "Name" ="MinOfPhysical"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Liquor"
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
    Bottom =129
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =23
        Top =8
        Right =119
        Bottom =122
        Top =0
        Name ="PalmOrderGuideFrom"
        Name =""
    End
    Begin
        Left =157
        Top =6
        Right =253
        Bottom =135
        Top =0
        Name ="Item"
        Name =""
    End
End
