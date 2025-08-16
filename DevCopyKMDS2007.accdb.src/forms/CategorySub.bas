Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =10
    ItemSuffix =11
    Top =600
    Bottom =600
    OrderBy ="CategorySub.Description"
    RecSrcDt = Begin
        0x8c1d7a589696e240
    End
    RecordSource ="SELECT SubCategory.* FROM SubCategory ORDER BY SubCategory.Description; "
    Caption ="SubCategory"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Rectangle
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Subform
            OldBorderStyle =0
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin Tab
            FontSize =9
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1020
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1620
                    Top =60
                    Width =3420
                    Height =360
                    ColumnWidth =2310
                    Name ="CategoryID"
                    ControlSource ="CategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Category.ID, Category.Description FROM Category ORDER BY Category.ID, Cat"
                        "egory.Description; "
                    ColumnWidths ="0;2880"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =60
                            Width =1440
                            Height =330
                            Name ="CategoryID_Label"
                            Caption ="CategoryID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1500
                    Top =540
                    Height =330
                    ColumnWidth =1440
                    TabIndex =1
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
                            OverlapFlags =93
                            Left =60
                            Top =540
                            Width =1440
                            Height =330
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =855
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1680
                    Top =60
                    Width =4020
                    Height =360
                    ColumnWidth =4185
                    Name ="Description"
                    ControlSource ="Description"
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
                            OverlapFlags =93
                            Left =240
                            Top =60
                            Width =1440
                            Height =330
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =540
                    TabIndex =1
                    Name ="TargetGP"
                    ControlSource ="TargetGP"
                    Format ="Percent"
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
                            TextAlign =0
                            Left =1020
                            Top =540
                            Width =900
                            Height =315
                            Name ="Label10"
                            Caption ="TargetGP:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =630
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6840
                    Top =300
                    Width =840
                    Height =330
                    ColumnWidth =1035
                    Name ="Change"
                    ControlSource ="Change"
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
                            OverlapFlags =93
                            Left =5400
                            Top =300
                            Width =1440
                            Height =330
                            Name ="Change_Label"
                            Caption ="Change"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1500
                    Width =1035
                    Height =330
                    ColumnWidth =1035
                    TabIndex =1
                    Name ="Setup"
                    ControlSource ="Setup"
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
                            OverlapFlags =93
                            Left =60
                            Width =1440
                            Height =330
                            Name ="Setup_Label"
                            Caption ="Setup"
                        End
                    End
                End
            End
        End
    End
End
