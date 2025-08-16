Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3960
    DatasheetFontHeight =10
    ItemSuffix =10
    Top =600
    Bottom =600
    Filter ="PriceGroupID<>0"
    RecSrcDt = Begin
        0x683f528c7497e240
    End
    RecordSource ="Item"
    Caption ="PriceGoupSub"
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
            Height =420
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1680
                    Width =945
                    Height =330
                    ColumnWidth =945
                    Name ="PriceGroupID"
                    ControlSource ="PriceGroupID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Width =1560
                            Height =330
                            Name ="PriceGroupID_Label"
                            Caption ="PriceGroupID"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1830
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1620
                    Top =120
                    Width =945
                    Height =330
                    ColumnWidth =945
                    ColumnOrder =0
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =120
                            Width =1560
                            Height =330
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1620
                    Top =960
                    Width =2310
                    Height =330
                    ColumnWidth =810
                    ColumnOrder =2
                    TabIndex =1
                    Name ="Price"
                    ControlSource ="Price"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =960
                            Width =1560
                            Height =330
                            Name ="Price_Label"
                            Caption ="Price"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =4320
                    Left =1680
                    Top =540
                    ColumnWidth =3075
                    ColumnOrder =1
                    TabIndex =2
                    Name ="cmbItem"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description, Item.Price, Item.PriceGroupID FROM Item; "
                    ColumnWidths ="0;1440;1440;1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =300
                            Top =540
                            Width =1245
                            Height =320
                            Name ="PriceGroupID1_Label"
                            Caption ="Description"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =390
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =2310
                    Height =330
                    ColumnWidth =2310
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =60
                            Width =1560
                            Height =330
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PriceGroupSub.cls"
