Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10590
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =570
    Top =1290
    Right =12945
    Bottom =9885
    RecSrcDt = Begin
        0x8b8d031b945fe340
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =6540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =1140
                    Width =3900
                    Height =5340
                    Name ="qryCashiersAndProgramsForm2"
                    SourceObject ="Form.CashierRightsProgramList"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =180
                            Width =3840
                            Height =840
                            Name ="Label0"
                            Caption ="Select one of the programs below then ...     Control access to the program with"
                                " the checkboxes on the left..."
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =4500
                    Top =1080
                    Width =5820
                    Height =5340
                    TabIndex =1
                    Name ="CashierRightsCashierListSub"
                    SourceObject ="Form.CashierRightsCashierList"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4560
                            Top =540
                            Width =5520
                            Height =480
                            FontSize =14
                            FontWeight =700
                            Name ="LabelMenuOption"
                            Caption ="LabelMenuOption"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    AccessKey =84
                    Left =4500
                    Top =60
                    Width =5520
                    Height =480
                    TabIndex =2
                    Name ="cmdSetAllowedForThisProgram"
                    Caption ="Se&t ALL Allowed Flags for PSC"
                    OnClick ="[Event Procedure]"
                    Tag ="PSC"
                    UnicodeAccessKey =116

                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "CashierRights.cls"
