Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =147
    Left =1170
    Top =2040
    RecSrcDt = Begin
        0xc4d73a329144e340
    End
    RecordSource ="qryItemSalesAndTransfersByMonth"
    Caption ="ItemSalesAndTransfersByMonth"
    OnOpen ="[Event Procedure]"
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
            ControlSource ="Location"
        End
        Begin BreakLevel
            ControlSource ="Description"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1680
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =300
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =10635
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Item Sales By Month"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1560
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =960
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Locations Between \" & Forms!PrintInventoryReport!cboLocationBegin & \" and \""
                        " & Forms!PrintInventoryReport!cboLocationEnd"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1200
                    Width =10621
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text116"
                    ControlSource ="=\"Vendors Between \" & Forms!PrintInventoryReport!cboVendorIDBegin & \" and \" "
                        "& Forms!PrintInventoryReport!cboVendorIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =3120
                    Top =120
                    Width =1215
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="TotalP_Label"
                    Caption ="On the shelf"
                    FontName ="Arial Narrow"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10635
                    Height =30
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10635
                    Height =30
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10635
                    Height =30
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10635
                    Height =30
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =60
                    Top =120
                    Width =960
                    Height =300
                    FontWeight =400
                    Name ="Label77"
                    Caption ="Location:"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =9420
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label79"
                    Caption ="Cur"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =1080
                    Top =120
                    Width =960
                    Height =300
                    FontWeight =400
                    Name ="Label85"
                    Caption ="Item"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =4380
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label117"
                    Caption ="Jan"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =4800
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label118"
                    Caption ="Feb"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =5220
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label119"
                    Caption ="Mar"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =5640
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label120"
                    Caption ="Apr"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =6060
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label121"
                    Caption ="May"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =6480
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label122"
                    Caption ="Jun"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =6900
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label123"
                    Caption ="Jul"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =7320
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label124"
                    Caption ="Aug"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =7740
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label125"
                    Caption ="Sep"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =8160
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label126"
                    Caption ="Oct"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =8580
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label127"
                    Caption ="Nov"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =34
                    Left =9000
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =400
                    Name ="Label128"
                    Caption ="Dec"
                    FontName ="Arial Narrow"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =900
                    Name ="Location"
                    ControlSource ="Location"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Top =60
                    Width =1020
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =1
                    Left =1020
                    Top =60
                    Width =2460
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =3
                    Left =4020
                    Top =60
                    Width =360
                    ColumnWidth =900
                    FontWeight =700
                    TabIndex =2
                    Name ="OnTheShelf"
                    ControlSource ="OnTheShelf"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =600
                    TabIndex =3
                    Name ="SizeD"
                    ControlSource ="SizeD"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =300
                End
                Begin Line
                    Left =60
                    Top =360
                    Width =10620
                    Name ="line5"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =4
                    Name ="CurrMM"
                    ControlSource ="CurrMMQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =5
                    Name ="JanQty"
                    ControlSource ="JanQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =6
                    Name ="FebQty"
                    ControlSource ="FebQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5280
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =7
                    Name ="MarQty"
                    ControlSource ="MarQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =8
                    Name ="AprQty"
                    ControlSource ="AprQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =9
                    Name ="MayQty"
                    ControlSource ="MayQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =10
                    Name ="JunQty"
                    ControlSource ="JunQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =11
                    Name ="JulQty"
                    ControlSource ="JulQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =12
                    Name ="AugQty"
                    ControlSource ="AugQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =13
                    Name ="SepQty"
                    ControlSource ="SepQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8220
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =14
                    Name ="OctQty"
                    ControlSource ="OctQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =15
                    Name ="NovQty"
                    ControlSource ="NovQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9060
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =16
                    Name ="DecQty"
                    ControlSource ="DecQty"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =465
                    Height =285
                    TabIndex =17
                    Name ="Text144"
                    ControlSource ="=[YtdQ]"
                    Format ="#;#;\"\";\"\""

                End
            End
        End
        Begin PageFooter
            Height =585
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =1920
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =4965
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =4740
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =4965
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =10680
                    Height =255
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Width =10605
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =960
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =420
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin TextBox
                    TextAlign =3
                    Left =4020
                    Top =60
                    Width =360
                    Name ="Text130"
                    ControlSource ="=Sum([OnTheShelf])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =1
                    Name ="Text131"
                    ControlSource ="=Sum([CurrMMQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =2
                    Name ="Text132"
                    ControlSource ="=Sum([JanQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =3
                    Name ="Text133"
                    ControlSource ="=Sum([FebQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5280
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =4
                    Name ="Text134"
                    ControlSource ="=Sum([MarQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =5
                    Name ="Text135"
                    ControlSource ="=Sum([AprQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =6
                    Name ="Text136"
                    ControlSource ="=Sum([MayQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =7
                    Name ="Text137"
                    ControlSource ="=Sum([JunQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =8
                    Name ="Text138"
                    ControlSource ="=Sum([JulQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =9
                    Name ="Text139"
                    ControlSource ="=Sum([AugQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =10
                    Name ="Text140"
                    ControlSource ="=Sum([SepQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8220
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =11
                    Name ="Text141"
                    ControlSource ="=Sum([OctQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =12
                    Name ="Text142"
                    ControlSource ="=Sum([NovQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9060
                    Top =60
                    Width =360
                    Height =255
                    TabIndex =13
                    Name ="Text143"
                    ControlSource ="=Sum([DecQty])"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =465
                    Height =285
                    TabIndex =14
                    Name ="Text145"
                    ControlSource ="=Sum([YtdQ])"
                    Format ="#;#;\"\";\"\""

                End
                Begin Label
                    Left =420
                    Top =450
                    Width =3030
                    Height =510
                    FontSize =10
                    ForeColor =0
                    Name ="Label146"
                    Caption ="Quantity = Sales - Counter + Club + Transfers reversed"
                End
            End
        End
    End
End
CodeBehindForm
' See "23 ItemSalesAndTransfersByMonth.cls"
