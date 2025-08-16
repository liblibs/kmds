Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
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
    ItemSuffix =64
    Left =2250
    Top =2280
    Right =13380
    Bottom =6285
    Filter ="((Not Item.OrderAt=0))"
    OrderBy ="[Lookup_cmbSizeID].[Description], Item.Description"
    RecSrcDt = Begin
        0x631e476053a0e240
    End
    RecordSource ="Item"
    Caption ="VendorItemsSub"
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
            Height =390
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =945
                    Height =330
                    ColumnWidth =945
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =60
                            Width =780
                            Height =330
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3120
                    Top =60
                    Width =945
                    Height =330
                    ColumnWidth =945
                    TabIndex =1
                    Name ="VendorID"
                    ControlSource ="VendorID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2340
                            Top =60
                            Width =780
                            Height =330
                            Name ="VendorID_Label"
                            Caption ="VendorID"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =3930
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1020
                    Top =660
                    Width =1140
                    Height =330
                    ColumnWidth =2550
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =660
                            Width =780
                            Height =330
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6540
                    Top =660
                    Width =945
                    Height =330
                    ColumnWidth =945
                    TabIndex =1
                    Name ="Onhand"
                    ControlSource ="Onhand"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5760
                            Top =660
                            Width =780
                            Height =330
                            Name ="Onhand_Label"
                            Caption ="Onhand"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4620
                    Top =720
                    Width =945
                    Height =330
                    ColumnWidth =855
                    TabIndex =2
                    Name ="OrderAt"
                    ControlSource ="OrderAt"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3840
                            Top =720
                            Width =780
                            Height =330
                            Name ="OrderAt_Label"
                            Caption ="OrderAt"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4740
                    Top =300
                    Width =945
                    Height =330
                    ColumnWidth =870
                    TabIndex =3
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =300
                            Width =780
                            Height =330
                            Name ="OrderTo_Label"
                            Caption ="OrderTo"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6000
                    Top =1200
                    Width =1140
                    Height =330
                    ColumnWidth =765
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5220
                            Top =1200
                            Width =780
                            Height =330
                            Name ="Cost_Label"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2820
                    Top =2220
                    Width =945
                    Height =330
                    ColumnWidth =1380
                    TabIndex =5
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2040
                            Top =2220
                            Width =780
                            Height =330
                            Name ="UnitsPerCase_Label"
                            Caption ="UnitsPerCase"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2820
                    Top =2640
                    Width =945
                    Height =330
                    ColumnWidth =1410
                    TabIndex =6
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2040
                            Top =2640
                            Width =780
                            Height =330
                            Name ="UnitsPerOrder_Label"
                            Caption ="UnitsPerOrder"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4800
                    Top =3060
                    Width =1140
                    Height =330
                    ColumnWidth =1365
                    TabIndex =7
                    Name ="StockNumber"
                    ControlSource ="StockNumber"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4020
                            Top =3060
                            Width =780
                            Height =330
                            Name ="StockNumber_Label"
                            Caption ="StockNumber"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2640
                    Top =1260
                    ColumnWidth =1020
                    TabIndex =8
                    Name ="cmbSizeID"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description FROM [Size] ORDER BY Size.ID; "
                    ColumnWidths ="0;1080"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =1200
                            Top =1260
                            Width =1155
                            Height =320
                            Name ="Description1_Label"
                            Caption ="Size"
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
' See "VendorItemsSub.cls"
