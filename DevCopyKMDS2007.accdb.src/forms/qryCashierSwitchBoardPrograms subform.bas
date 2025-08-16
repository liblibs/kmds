Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    DatasheetFontHeight =10
    ItemSuffix =14
    RecSrcDt = Begin
        0xd90e6ecaf79de240
    End
    RecordSource ="qryCashierSwitchBoardPrograms"
    Caption ="qryCashierSwitchBoardPrograms subform"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2835
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =600
                    Height =255
                    ColumnWidth =600
                    Name ="ID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =480
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Name"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="Name_Label"
                            Caption ="Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =840
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Type"
                    ControlSource ="Type"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1560
                            Height =255
                            Name ="Type_Label"
                            Caption ="Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1200
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =3
                    Name ="Allowed"
                    ControlSource ="Allowed"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1560
                            Height =255
                            Name ="Allowed_Label"
                            Caption ="Allowed"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1560
                    Width =4560
                    Height =450
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="Argument"
                    ControlSource ="Argument"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1560
                            Height =255
                            Name ="Argument_Label"
                            Caption ="Argument"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2100
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =5
                    Name ="SwitchboardID"
                    ControlSource ="SwitchboardID"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2100
                            Width =1560
                            Height =255
                            Name ="SwitchboardID_Label"
                            Caption ="SwitchboardID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2460
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    Name ="ItemNumber"
                    ControlSource ="ItemNumber"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =1560
                            Height =255
                            Name ="ItemNumber_Label"
                            Caption ="ItemNumber"
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
