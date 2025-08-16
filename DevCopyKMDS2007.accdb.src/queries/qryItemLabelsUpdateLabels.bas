Operation =4
Option =0
Where ="(((ItemLabels.ID)=[parItemID]))"
Begin InputTables
    Name ="ItemLabels"
End
Begin OutputColumns
    Name ="ItemLabels.Labels"
    Expression ="1"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
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
        Name ="ItemLabels"
        Name =""
    End
End
