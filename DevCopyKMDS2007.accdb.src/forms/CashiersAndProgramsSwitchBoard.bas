Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12090
    DatasheetFontHeight =10
    ItemSuffix =15
    RecSrcDt = Begin
        0x7e298cdcf79de240
    End
    RecordSource ="qryCashierSwitchBoardMain"
    Caption ="CashiersAndProgramsSwitchBoard"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1440
                    Top =120
                    Name ="cmbCashier"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Cashier.ID, Cashier.Name FROM Cashier ORDER BY [Name]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =120
                            Width =1110
                            Height =245
                            Name ="Cashier Name_Label"
                            Caption ="Cashier Name"
                            EventProcPrefix ="Cashier_Name_Label"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3840
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =420
                    Top =300
                    Width =3975
                    Height =2580
                    Name ="List10"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryCashierSwitchBoardMain.ItemText, qryCashierSwitchBoardMain.ItemNumber "
                        "FROM qryCashierSwitchBoardMain; "
                    ColumnWidths ="2535;1440"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Subform
                    OverlapFlags =85
                    Left =4740
                    Top =300
                    Width =6870
                    Height =3120
                    TabIndex =1
                    Name ="CashiersAndProgramsSwitchBoardSub"
                    SourceObject ="Form.CashiersAndProgramsSwitchBoardSub"
                    LinkChildFields ="SwitchboardID"
                    LinkMasterFields ="Arg"

                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "CashiersAndProgramsSwitchBoard.cls"
