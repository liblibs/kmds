Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =137
    Left =585
    Top =1320
    OrderBy ="SumOfPrice DESC"
    RecSrcDt = Begin
        0xe792d20b39fce240
    End
    RecordSource ="qryItemSalesByCategoryReport"
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
            GroupFooter = NotDefault
            ControlSource ="SalesDptID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Category"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2100
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-15
                    Top =60
                    Width =10755
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Item Sales by Category"
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
                    Top =1140
                    Width =10680
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1200
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Sales based on Sales Between \" & Format(Forms!PrintInventoryReport!txtDateRa"
                        "ngeFrom,\"mm/dd/yyyy\") & \" And \" & Format(Forms!PrintInventoryReport!txtDateR"
                        "angeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1500
                    Width =10621
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1800
                    Width =10621
                    Height =300
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text125"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboCategoryIDBegin & \" and \" & Form"
                        "s!PrintInventoryReport!cboCategoryIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =795
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =180
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7800
                    Top =120
                    Width =540
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10620
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6660
                    Top =120
                    Width =840
                    Height =300
                    Name ="Label86"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9660
                    Top =120
                    Width =1020
                    Height =600
                    Name ="Label87"
                    Caption ="Dollars Sold"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =780
                    Width =10560
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8400
                    Top =120
                    Width =1020
                    Height =585
                    Name ="Label113"
                    Caption ="Quantity Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5475
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Label118"
                    Caption ="Last Sold"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Top =480
                    Width =1155
                    Name ="txtSetupAfter"
                    ControlSource ="=Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\")"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            TextAlign =2
                            Left =4230
                            Top =120
                            Width =1215
                            Height =300
                            Name ="Label117"
                            Caption ="Setup After"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    ColumnWidth =615
                    Name ="SaleDpt"
                    ControlSource ="SalesDptID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =3060
                    ColumnWidth =2730
                    TabIndex =1
                    Name ="SaleDptDesc"
                    ControlSource ="SalesDptDesc"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =1260
                    Width =2394
                    ColumnWidth =1185
                    FontWeight =700
                    Name ="GroupDesc"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =3654
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    Left =60
                    Width =1140
                    FontWeight =700
                    TabIndex =1
                    Name ="Group"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Width =720
                    Height =225
                    ColumnWidth =885
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6780
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Width =1260
                    ColumnWidth =1770
                    TabIndex =4
                    Name ="SumOfPrice"
                    ControlSource ="SumOfPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8280
                    Width =1080
                    TabIndex =5
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4620
                    Width =1080
                    TabIndex =6
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="Short Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5835
                    Width =900
                    TabIndex =7
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="Short Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5835
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3720
                    Width =840
                    TabIndex =8
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =240
                End
                Begin Line
                    Visible = NotDefault
                    Left =3600
                    Top =240
                    Width =7080
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =1260
                    Name ="Text130"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            Left =7620
                            Top =60
                            Width =540
                            Height =300
                            Name ="Label131"
                            Caption ="Total"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8280
                    Top =60
                    Width =1080
                    TabIndex =1
                    Name ="Text132"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4080
                    Top =60
                    Width =3480
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Category"
                    ControlSource ="Category"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =60
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =4500
                    Top =60
                    Width =6120
                    Name ="Line136"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =240
                    Width =1260
                    Name ="Text127"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =7620
                            Top =180
                            Width =540
                            Height =300
                            Name ="Label128"
                            Caption ="Total"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8280
                    Top =240
                    Width =1080
                    TabIndex =1
                    Name ="Text129"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2760
                    Top =180
                    Width =4800
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text133"
                    ControlSource ="SalesDptDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =180
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =480
                End
                Begin Line
                    Left =4560
                    Top =60
                    Width =6120
                    Name ="Line134"
                End
                Begin Line
                    Left =4560
                    Top =120
                    Width =6120
                    Name ="Line135"
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
                    Width =4380
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
                    Width =10620
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
            Height =480
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Top =120
                    Width =1260
                    Name ="Text122"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            Left =4980
                            Top =120
                            Width =2040
                            Height =300
                            Name ="Label124"
                            Caption ="Report Total"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8220
                    Top =120
                    Width =1080
                    TabIndex =1
                    Name ="Text123"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                End
            End
        End
    End
End
CodeBehindForm
' See "43 Item Sales By Category.cls"
