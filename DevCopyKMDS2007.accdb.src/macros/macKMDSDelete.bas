Version =196611
ColumnsShown =10
Begin
    Action ="OnError"
    Argument ="0"
End
Begin
    Action ="GoToControl"
    Argument ="=[Screen].[PreviousControl].[Name]"
End
Begin
    Action ="ClearMacroError"
End
Begin
    Condition ="Not [Form].[NewRecord]"
    Action ="RunCommand"
    Argument ="223"
End
Begin
    Condition ="[Form].[NewRecord] And Not [Form].[Dirty]"
    Action ="Beep"
End
Begin
    Condition ="[Form].[NewRecord] And [Form].[Dirty]"
    Action ="RunCommand"
    Argument ="292"
End
Begin
    Condition ="[MacroError]<>0"
    Action ="MsgBox"
    Argument ="=[MacroError].[Description]"
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="RunCode"
    Argument =" SetRunMailChangesTrue () "
End
