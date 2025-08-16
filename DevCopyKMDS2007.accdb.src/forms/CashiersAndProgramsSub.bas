Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =9
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4554
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =2505
    Top =2985
    Right =8325
    Bottom =7500
    HelpContextId =7300
    OnUnload ="[Event Procedure]"
    Filter ="([CashiersAndPrograms].[CashierID]=66)"
    OrderBy ="[CashiersAndPrograms].[CashierID]"
    RecSrcDt = Begin
        0xacc5ce56cf38e440
    End
    RecordSource ="CashiersAndPrograms"
    Caption ="CashiersAndPrograms"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            SpecialEffect =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    Left =1920
                    Width =2490
                    ColumnWidth =3120
                    ColumnOrder =0
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Cashier.ID, Cashier.Name FROM Cashier; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="CashierID"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1260
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =85
                    Left =1920
                    Top =840
                    ColumnWidth =1650
                    Name ="Allowed"
                    ControlSource ="Allowed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =840
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="Allowed"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =2100
                    Top =300
                    Width =1980
                    ColumnWidth =3525
                    TabIndex =1
                    BoundColumn =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cmbProgramName"
                    ControlSource ="ProgramName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [FirstLetter] & \" \" & [Switchboard Items].ItemText AS IT, [Switchboard "
                        "Items].Argument, [Switchboard Items].SwitchboardID, [Switchboard Items].ItemNumb"
                        "er, qrySwitchBoardFirstLetter.FirstLetter, [Switchboard Items].ItemText FROM [Sw"
                        "itchboard Items] INNER JOIN qrySwitchBoardFirstLetter ON [Switchboard Items].Swi"
                        "tchboardID=qrySwitchBoardFirstLetter.SwitchboardID ORDER BY [Switchboard Items]."
                        "SwitchboardID, [Switchboard Items].ItemNumber; "
                    ColumnWidths ="1440;2880"
                    FontName ="MS Sans Serif"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =855
                            Top =300
                            Width =1185
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="Program Name:"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =840
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1920
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Width =840
                            Height =240
                            Name ="Label7"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1920
                    Top =300
                    TabIndex =1
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =300
                            Width =975
                            Height =240
                            Name ="Label9"
                            Caption ="ChangeDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1800
                    Top =600
                    Width =2490
                    ColumnWidth =2715
                    TabIndex =2
                    Name ="ProgramName"
                    ControlSource ="ProgramName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =600
                            Width =1080
                            Height =240
                            Name ="Label3"
                            Caption ="ProgramName"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CashiersAndProgramsSub.cls"
