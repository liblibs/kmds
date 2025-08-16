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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =131
    Left =1560
    Top =4650
    OrderBy ="SumOfPrice DESC"
    RecSrcDt = Begin
        0xfe206a0bd713e340
    End
    RecordSource ="qryItemSalesReport"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1800
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
                    Left =-45
                    Top =60
                    Width =10815
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Item Sales"
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
                    Left =7860
                    Top =120
                    Width =540
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =120
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =420
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
                    Left =6960
                    Top =120
                    Width =840
                    Height =300
                    Name ="Label86"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6960
                    LayoutCachedTop =120
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9540
                    Top =120
                    Width =1020
                    Height =600
                    Name ="Label87"
                    Caption ="Dollars Sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9540
                    LayoutCachedTop =120
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =720
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
                    Left =8580
                    Top =120
                    Width =1020
                    Height =585
                    Name ="Label113"
                    Caption ="Quantity Sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =705
                End
                Begin Label
                    Left =5955
                    Top =120
                    Width =1005
                    Height =300
                    Name ="Label118"
                    Caption ="Last Sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5955
                    LayoutCachedTop =120
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4740
                    Top =480
                    Width =1155
                    Name ="txtSetupAfter"
                    ControlSource ="=Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\")"
                    Format ="Short Date"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =480
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            TextAlign =2
                            Left =4710
                            Top =120
                            Width =1215
                            Height =300
                            Name ="Label117"
                            Caption ="Setup after"
                            Tag ="DetachedLabel"
                            LayoutCachedLeft =4710
                            LayoutCachedTop =120
                            LayoutCachedWidth =5925
                            LayoutCachedHeight =420
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
        Begin Section
            KeepTogether = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =1380
                    Width =2394
                    ColumnWidth =1185
                    FontWeight =700
                    Name ="GroupDesc"
                    ControlSource ="Description"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =3774
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    Left =45
                    Width =1275
                    FontWeight =700
                    TabIndex =1
                    Name ="Group"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7800
                    Width =720
                    Height =225
                    ColumnWidth =885
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="Price"

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7080
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Onhand"
                    ControlSource ="Onhand"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Width =900
                    ColumnWidth =1770
                    TabIndex =4
                    Name ="SumOfPrice"
                    ControlSource ="SumOfPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8820
                    Width =780
                    TabIndex =5
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"
                    Format ="Standard"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4620
                    Width =1245
                    TabIndex =6
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="Short Date"

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =5865
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5835
                    Width =1245
                    TabIndex =7
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="Short Date"

                    LayoutCachedLeft =5835
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3840
                    Width =720
                    TabIndex =8
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"

                    LayoutCachedLeft =3840
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
            Name ="GroupFooter4"
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
            Height =840
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Top =120
                    Name ="Text122"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =120
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            Left =4740
                            Top =120
                            Width =2040
                            Height =300
                            Name ="Label124"
                            Caption ="Report Total"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =120
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7740
                    Top =120
                    Width =1320
                    TabIndex =1
                    Name ="Text123"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =120
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =360
                End
            End
        End
    End
End
CodeBehindForm
' See "12 Item Sales.cls"
