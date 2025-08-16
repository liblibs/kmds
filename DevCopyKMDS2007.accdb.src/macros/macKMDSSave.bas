Version =196611
ColumnsShown =10
Begin
    Action ="OnError"
    Argument ="0"
End
Begin
    Action ="RunCommand"
    Argument ="97"
End
Begin
    Condition ="[MacroError]<>0"
    Action ="MsgBox"
    Argument ="=[MacroError].[Description]"
    Argument ="-1"
    Argument ="0"
End
