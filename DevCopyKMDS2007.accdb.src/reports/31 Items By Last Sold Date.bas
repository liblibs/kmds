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
    ItemSuffix =80
    Left =1170
    Top =2040
    OrderBy ="ID"
    RecSrcDt = Begin
        0xe8adf647b2cae240
    End
    RecordSource ="qryInventoryValueByCodeByDateLastSold"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1620
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =6105
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Report by Date Last Sold"
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
                    Top =1140
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Top =600
                    Width =10621
                    Height =480
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text84"
                    ControlSource ="=\"Date Last Sold before \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom"
                        ",\"mm/dd/yyyy\")"
                    Format ="Long Date"
                    FontName ="Times New Roman"

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
                    Left =5490
                    Top =120
                    Width =615
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6330
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
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =1
                    Left =60
                    Width =1260
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1380
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =4260
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
                    Format ="#,##0;-#,##0[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =7380
                    Width =300
                    Height =225
                    TabIndex =3
                    Name ="SalePriceItem"
                    ControlSource ="=IIf(IsNull([SalePrice]),\"\",\"*\")"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =4860
                    Width =1200
                    Height =225
                    TabIndex =4
                    Name ="CurPrice"
                    ControlSource ="=IIf(IsNull([SalePrice]),[Price],[SalePrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Width =1200
                    Height =225
                    TabIndex =5
                    Name ="ExtCurPrice"
                    ControlSource ="=[CurPrice]*[Onhand]"
                    Format ="#,##0.00;-#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =1
                    Left =7800
                    Width =540
                    Height =225
                    TabIndex =6
                    Name ="Text51"
                    ControlSource ="=IIf([CurPrice]=0,0,([CurPrice]-[Cost])/[CurPrice])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Visible = NotDefault
                    Left =240
                    Top =240
                    Width =10380
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8460
                    Width =2160
                    ColumnWidth =2400
                    TabIndex =7
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =120
                    Width =1200
                    Height =225
                    TabIndex =1
                    Name ="txtDepartmentRetail"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2040
                    Top =120
                    Width =420
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="Department"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =60
                    Top =120
                    Width =1920
                    Height =225
                    TabIndex =3
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =7560
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text52"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    Left =8340
                    Top =120
                    Width =660
                    Height =225
                    TabIndex =5
                    Name ="TargetGP"
                    ControlSource ="TargetGP"
                    Format ="Percent"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"

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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =120
                    Width =1200
                    Height =225
                    TabIndex =1
                    Name ="Text38"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =960
                    Top =120
                    Width =2880
                    Height =225
                    TabIndex =2
                    Name ="Text47"
                    ControlSource ="GroupDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =7560
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Text53"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"

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
                    TextFontFamily =18
                    Left =3840
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =60
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text43"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Left =600
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =9360
                    Top =60
                    Width =720
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text54"
                    ControlSource ="=(Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])-Sum([Onhand]*[Cost]"
                        "))/Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Percent"
                    FontName ="Times New Roman"

                End
            End
        End
    End
End
CodeBehindForm
' See "31 Items By Last Sold Date.cls"
