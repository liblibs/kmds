Operation =4
Option =0
Begin InputTables
    Name ="OLCCWebItems"
    Name ="SubCategory"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.SubCategoryID"
    Expression ="SubCategory.id"
End
Begin Joins
    LeftTable ="OLCCWebItems"
    RightTable ="SubCategory"
    Expression ="OLCCWebItems.Category = SubCategory.Description"
    Flag =1
    LeftTable ="OLCCWebItems"
    RightTable ="Item"
    Expression ="OLCCWebItems.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
        Name ="OLCCWebItems"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="SubCategory"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
