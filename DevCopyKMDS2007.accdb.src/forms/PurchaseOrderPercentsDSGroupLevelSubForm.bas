Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2775
    DatasheetFontHeight =10
    ItemSuffix =12
    RecSrcDt = Begin
        0x7e0ccd1eb06be340
    End
    RecordSource ="SELECT DaysSupplyGroupLevels.GroupLevel, DaysSupplyGroupLevels.DaysSupply, DaysS"
        "upplyGroupLevels.MinUnitSales, DaysSupplyGroupLevels.MaxUnitSales, DaysSupplyGro"
        "upLevels.SetupDate, DaysSupplyGroupLevels.ChangeDate FROM DaysSupplyGroupLevels "
        "ORDER BY DaysSupplyGroupLevels.GroupLevel; "
    Caption ="PurchaseOrderPercentsDSGroupLevelSubForm"
    DatasheetFontName ="Trebuchet MS"
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
            Height =2295
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =1230
                    ColumnOrder =0
                    Name ="GroupLevel"
                    ControlSource ="GroupLevel"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="GroupLevel_Label"
                            Caption ="Group Level"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =480
                    Width =900
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =3
                    TabIndex =1
                    Name ="DaysSupply"
                    ControlSource ="DaysSupply"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="DaysSupply_Label"
                            Caption ="Days Supply"
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
                    ColumnWidth =1425
                    ColumnOrder =1
                    TabIndex =2
                    Name ="MinUnitSales"
                    ControlSource ="MinUnitSales"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1560
                            Height =255
                            Name ="MinUnitSales_Label"
                            Caption ="Min Unit Sales"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1200
                    Width =900
                    Height =255
                    ColumnWidth =1470
                    ColumnOrder =2
                    TabIndex =3
                    Name ="MaxUnitSales"
                    ControlSource ="MaxUnitSales"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1560
                            Height =255
                            Name ="MaxUnitSales_Label"
                            Caption ="Max Unit Sales"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1560
                    Width =1035
                    Height =255
                    ColumnWidth =2070
                    TabIndex =4
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1560
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1920
                    Width =1035
                    Height =255
                    ColumnWidth =2070
                    TabIndex =5
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1560
                            Height =255
                            Name ="ChangeDate_Label"
                            Caption ="Changed"
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
