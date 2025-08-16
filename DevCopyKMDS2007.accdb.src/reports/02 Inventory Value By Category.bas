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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =80
    Left =585
    Top =1320
    OrderBy ="ID"
    RecSrcDt = Begin
        0xe04a769c29dfe240
    End
    RecordSource ="qryInventoryValueByCategory"
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
            ControlSource ="Group"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Department"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CategoryID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SubCategoryDescription"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1335
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =5040
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Value Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Top =660
                    Width =10561
                    Height =300
                    ColumnOrder =0
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Top =960
                    Width =10561
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboCategoryIDBegin & \" and \" & Form"
                        "s!PrintInventoryReport!cboCategoryIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =495
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =240
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3990
                    Top =120
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    Left =5100
                    Top =120
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5790
                    Top =120
                    Width =855
                    Height =300
                    Name ="D2_Label"
                    Caption ="Ext Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7290
                    Top =120
                    Width =615
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8130
                    Top =120
                    Width =1050
                    Height =300
                    Name ="D4_Label"
                    Caption =" Extended"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =420
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    Width =840
                    Height =420
                    FontSize =18
                    ForeColor =8388608
                    Name ="Group"
                    ControlSource ="Group"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    Left =720
                    Width =7119
                    Height =420
                    ColumnWidth =1185
                    FontSize =18
                    TabIndex =1
                    ForeColor =8388608
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =375
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =180
                    Width =672
                    Height =375
                    ColumnWidth =1260
                    FontSize =14
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =900
                    Width =2304
                    Height =375
                    ColumnWidth =1710
                    FontSize =14
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =330
            BreakLevel =2
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =420
                    Width =6051
                    Height =330
                    ColumnWidth =4170
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="CategoryDescription"
                    ControlSource ="CategoryDescription"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            BreakLevel =3
            Name ="GroupHeader4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =120
                    Width =5763
                    Height =300
                    ColumnWidth =4170
                    FontSize =11
                    ForeColor =8388608
                    Name ="SubCategoryDescription"
                    ControlSource ="SubCategoryDescription"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =180
                    Width =1020
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =180
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1260
                    Width =3120
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =4020
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Width =900
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="ExtCost"
                    ControlSource ="=[Onhand]*[Cost]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =4800
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4800
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =9180
                    Width =300
                    Height =225
                    TabIndex =5
                    Name ="SalePriceItem"
                    ControlSource ="=IIf(IsNull([SalePrice]),\"\",\"*\")"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9180
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =6660
                    Width =1200
                    Height =225
                    TabIndex =6
                    Name ="CurPrice"
                    ControlSource ="=IIf(IsNull([SalePrice]),[Price],[SalePrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7920
                    Width =1200
                    Height =225
                    TabIndex =7
                    Name ="ExtCurPrice"
                    ControlSource ="=[CurPrice]*[Onhand]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =1
                    Left =9600
                    Width =540
                    Height =225
                    TabIndex =8
                    Name ="Text51"
                    ControlSource ="=IIf([CurPrice]=0,0,([CurPrice]-[Cost])/[CurPrice])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =240
                    Top =300
                    Width =10380
                    Name ="Line5"
                    LayoutCachedLeft =240
                    LayoutCachedTop =300
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =300
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter5"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =2040
                    Height =225
                    Name ="Text59"
                    ControlSource ="SubCategoryDescription"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4020
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text60"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Top =60
                    Width =900
                    Height =225
                    TabIndex =2
                    Name ="Text61"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7920
                    Top =60
                    Width =1200
                    Height =225
                    TabIndex =3
                    Name ="Text62"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9360
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text66"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    Left =10140
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =5
                    Name ="Text67"
                    ControlSource ="TargetGP"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =300
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1020
                    Width =1740
                    Height =225
                    Name ="Text58"
                    ControlSource ="CategoryDescription"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4020
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text68"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Top =60
                    Width =900
                    Height =225
                    TabIndex =2
                    Name ="Text69"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7920
                    Top =60
                    Width =1200
                    Height =225
                    TabIndex =3
                    Name ="Text70"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9300
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text71"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    Left =10080
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =5
                    Name ="Text72"
                    ControlSource ="TargetGP"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4020
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text26"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="txtDepartmentCost"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7920
                    Top =120
                    Width =1200
                    Height =225
                    TabIndex =2
                    Name ="txtDepartmentRetail"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2040
                    Top =120
                    Width =420
                    Height =225
                    TabIndex =3
                    Name ="Text34"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =60
                    Top =120
                    Width =1920
                    Height =225
                    TabIndex =4
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9360
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text52"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    Left =10140
                    Top =120
                    Width =660
                    Height =225
                    TabIndex =6
                    Name ="TargetGP"
                    ControlSource ="TargetGP"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4020
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text36"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7920
                    Top =120
                    Width =1200
                    Height =225
                    TabIndex =2
                    Name ="Text38"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =960
                    Top =120
                    Width =2880
                    Height =225
                    TabIndex =3
                    Name ="Text47"
                    ControlSource ="GroupDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9360
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text53"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin PageFooter
            Height =555
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
                    Width =5040
                    Height =255
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
                    Width =4260
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
                    Width =5040
                    Height =255
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
                    Left =120
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
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3840
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5520
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7800
                    Top =60
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text43"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin Label
                    TextFontFamily =34
                    Left =600
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9360
                    Top =60
                    Width =720
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text54"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "02 Inventory Value By Category.cls"
