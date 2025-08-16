Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4050
    DatasheetFontHeight =10
    ItemSuffix =9
    HelpContextId =7300
    OnUnload ="[Event Procedure]"
    Filter ="((Lookup_ProgramName.ItemText='4. PSC Functions'))"
    RecSrcDt = Begin
        0xae59065d0a9ee240
    End
    RecordSource ="CashiersAndPrograms"
    Caption ="CashiersAndProgramsDatasheet"
    AfterUpdate ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            Height =255
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =1215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =25
                    Left =1680
                    Top =480
                    Width =2310
                    Height =255
                    ColumnWidth =2220
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    RowSourceType ="Table/Query"
                    RowSource ="Cashier"
                    ColumnWidths ="0;2880"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="CashierID_Label"
                            Caption ="CashierID"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1680
                    Top =840
                    Width =2310
                    Height =255
                    ColumnWidth =0
                    TabIndex =2
                    BoundColumn =1
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="ProgramName"
                    ControlSource ="ProgramName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Switchboard Items].ItemText, [Switchboard Items].Argument, [Switchboard "
                        "Items].SwitchboardID, [Switchboard Items].ItemNumber FROM [Switchboard Items] OR"
                        "DER BY [Switchboard Items].SwitchboardID, [Switchboard Items].ItemNumber; "
                    ColumnWidths ="1440;2880"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =840
                            Width =1560
                            Height =255
                            Name ="ProgramName_Label"
                            Caption ="ProgramName"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1620
                    Top =150
                    ColumnWidth =855
                    Name ="chkAllowed"
                    ControlSource ="Allowed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =900
                            Top =120
                            Width =630
                            Height =240
                            Name ="Label7"
                            Caption ="Allowed"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "CashiersAndProgramsDatasheet.cls"
