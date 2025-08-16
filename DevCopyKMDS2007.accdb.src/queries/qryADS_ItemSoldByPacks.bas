Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\") AND ((InvoLine.Setup"
    "Date)>Now()-365))"
Having ="(((Item.UnitsPerPack)>0))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.UnitsPerPack"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.SubPack = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="InvoLine"
    Expression ="UPCItem.ItemID = InvoLine.Code"
    Flag =1
End
Begin Groups
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.UnitsPerPack"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "OrderBy" ="qryADS_ItemSoldByPacks.ID"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SumOfQuantity"
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
        Left =242
        Top =9
        Right =338
        Bottom =123
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =469
        Top =-5
        Right =565
        Bottom =109
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =618
        Top =-7
        Right =714
        Bottom =107
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
