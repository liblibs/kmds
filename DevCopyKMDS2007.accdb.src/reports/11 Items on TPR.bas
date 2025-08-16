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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =79
    Left =585
    Top =1320
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xd0c19638f667e340
    End
    RecordSource ="qryInventoryValueByCategoryOnTPR"
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
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10681
                    Height =420
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
                    Width =10695
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items on TPR"
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
                    Left =4200
                    Top =120
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =120
                    LayoutCachedWidth =5025
                    LayoutCachedHeight =420
                End
                Begin Label
                    Visible = NotDefault
                    Left =5310
                    Top =120
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5310
                    LayoutCachedTop =120
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =6000
                    Top =120
                    Width =855
                    Height =300
                    Name ="D2_Label"
                    Caption ="Ext Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =120
                    LayoutCachedWidth =6855
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =7500
                    Top =120
                    Width =615
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7500
                    LayoutCachedTop =120
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =8340
                    Top =120
                    Width =1050
                    Height =300
                    Name ="D4_Label"
                    Caption =" Extended"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =120
                    LayoutCachedWidth =9390
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =150
                    LayoutCachedWidth =10770
                    LayoutCachedHeight =150
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
                    Left =60
                    Width =1320
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1440
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =4320
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6000
                    Width =900
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="ExtCost"
                    ControlSource ="=[Onhand]*[Cost]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6000
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =5100
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =9480
                    Width =300
                    Height =225
                    TabIndex =5
                    Name ="SalePriceItem"
                    ControlSource ="=IIf(IsNull([SalePrice]),\"\",\"*\")"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9480
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =6960
                    Width =1200
                    Height =225
                    TabIndex =6
                    Name ="CurPrice"
                    ControlSource ="=IIf(IsNull([SalePrice]),[Price],[SalePrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6960
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8220
                    Width =1200
                    Height =225
                    TabIndex =7
                    Name ="ExtCurPrice"
                    ControlSource ="=[CurPrice]*[Onhand]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =1
                    Left =9900
                    Width =540
                    Height =225
                    TabIndex =8
                    Name ="Text51"
                    ControlSource ="=IIf([CurPrice]=0,0,([CurPrice]-[Cost])/[CurPrice])"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9900
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =240
                    Width =9960
                    Name ="Line5"
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
                    Left =5700
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="txtDepartmentCost"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

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
                    Left =5700
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

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
                    Width =5040
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
                    Left =60
                    Width =10680
                    Height =255
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
                    Left =5520
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
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
                    TabIndex =2
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
                    TabIndex =3
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
' See "11 Items on TPR.cls"
