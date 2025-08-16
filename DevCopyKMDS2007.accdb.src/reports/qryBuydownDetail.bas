Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =40
    RecSrcDt = Begin
        0xe38da5a5cdd3e240
    End
    RecordSource ="qryBuydownDetail"
    Caption ="qryBuydownDetail"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="RepsName"
        End
        Begin BreakLevel
            ControlSource ="Description"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =3165
                    Height =510
                    FontSize =20
                    Name ="Label32"
                    Caption ="qryBuydownDetail"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line35"
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1140
                    Height =300
                    Name ="RepsName_Label"
                    Caption ="RepsName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1260
                    Top =60
                    Width =960
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2280
                    Top =60
                    Width =840
                    Height =300
                    Name ="ReferenceNo_Label"
                    Caption ="ReferenceNo"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3180
                    Top =60
                    Width =480
                    Height =300
                    Name ="EndDate_Label"
                    Caption ="EndDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3720
                    Top =60
                    Width =480
                    Height =300
                    Name ="BeginDate_Label"
                    Caption ="BeginDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4260
                    Top =60
                    Width =540
                    Height =300
                    Name ="BuydownQty_Label"
                    Caption ="BuydownQty"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4860
                    Top =60
                    Width =960
                    Height =300
                    Name ="BuydownDollars_Label"
                    Caption ="BuydownDollars"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =960
                    Height =300
                    Name ="DollarsToDate_Label"
                    Caption ="DollarsToDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6900
                    Top =60
                    Width =480
                    Height =300
                    Name ="Quantity_Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =9360
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =9360
                    Name ="Line37"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =960
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =1140
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="RepsName"
                    ControlSource ="RepsName"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1320
                    Width =960
                    Height =270
                    ColumnWidth =3150
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2400
                    Width =840
                    Height =270
                    TabIndex =2
                    Name ="ReferenceNo"
                    ControlSource ="ReferenceNo"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3300
                    Top =360
                    Width =1560
                    Height =270
                    TabIndex =3
                    Name ="EndDate"
                    ControlSource ="EndDate"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1320
                    Top =360
                    Width =1860
                    Height =270
                    ColumnWidth =2145
                    TabIndex =4
                    Name ="BeginDate"
                    ControlSource ="BeginDate"

                End
                Begin Label
                    TextAlign =3
                    Left =3240
                    Top =660
                    Width =480
                    Height =300
                    Name ="ItemsCode_Label"
                    Caption ="ItemsCode"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5280
                    Top =660
                    Width =1335
                    Height =300
                    Name ="QuantitySold_Label"
                    Caption ="Quantity Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7080
                    Top =660
                    Width =720
                    Height =300
                    Name ="DiscountDollars_Label"
                    Caption ="DiscountDollars"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =3180
                    Width =720
                    Height =225
                    Name ="ItemsCode"
                    ControlSource ="ItemsCode"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Width =840
                    Height =225
                    TabIndex =1
                    Name ="QuantitySold"
                    ControlSource ="QuantitySold"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7080
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="DiscountDollars"
                    ControlSource ="DiscountDollars"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4140
                    Width =1560
                    Height =225
                    ColumnWidth =3150
                    TabIndex =3
                    Name ="Text38"
                    ControlSource ="Description"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =630
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =60
                    Width =9240
                    Name ="Text24"
                    ControlSource ="=\"Summary for \" & \"'RepsName' = \" & \" \" & [RepsName] & \" (\" & Count(*) &"
                        " \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =285
                    Width =465
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label25"
                    Caption ="Sum"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Top =300
                    Width =660
                    Height =285
                    TabIndex =1
                    Name ="Sum Of QuantitySold"
                    ControlSource ="=Sum([QuantitySold])"
                    EventProcPrefix ="Sum_Of_QuantitySold"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =300
                    Width =1140
                    Height =285
                    TabIndex =2
                    Name ="Sum Of DiscountDollars"
                    ControlSource ="=Sum([DiscountDollars])"
                    Format ="$#,##0.00;($#,##0.00)"
                    EventProcPrefix ="Sum_Of_DiscountDollars"

                End
            End
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text33"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =60
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text34"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
