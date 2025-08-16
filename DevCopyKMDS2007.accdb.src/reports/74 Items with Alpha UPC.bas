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
    ItemSuffix =90
    Left =1710
    Top =720
    RecSrcDt = Begin
        0x1c09834ceabce440
    End
    RecordSource ="qryItemsWithAlphaUPC"
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
                    Top =660
                    Width =10681
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =10741
                    LayoutCachedHeight =1080
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
                    Caption ="Items with Alpha UPC"
                    LayoutCachedLeft =45
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =570
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
            Height =840
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =3240
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =120
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =6375
                    Top =120
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6375
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =420
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    Left =7440
                    Top =120
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =120
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =8415
                    Top =180
                    Width =600
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8415
                    LayoutCachedTop =180
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =480
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
                Begin Label
                    TextAlign =3
                    Left =9720
                    Top =120
                    Width =1050
                    Height =300
                    Name ="Label85"
                    Caption ="Date sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9720
                    LayoutCachedTop =120
                    LayoutCachedWidth =10770
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =1305
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Label88"
                    Caption ="SN"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1305
                    LayoutCachedTop =120
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =120
                    Top =120
                    Width =1185
                    Height =300
                    Name ="Label89"
                    Caption ="Alpha Code"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =420
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
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =2340
                    Width =1020
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =3420
                    Width =3060
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =6480
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6480
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7140
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7140
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8220
                    Width =795
                    Height =225
                    TabIndex =4
                    Name ="CurPrice"
                    ControlSource ="Price"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =240
                    Width =9960
                    Name ="Line5"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9120
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text80"
                    ControlSource ="=IIf([price]>0,(([Price]-[Cost])/[Price]),0)"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9120
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =120
                    Width =1140
                    Height =225
                    TabIndex =6
                    Name ="Text81"
                    ControlSource ="AlphaUPC"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9840
                    Width =960
                    Height =225
                    TabIndex =7
                    Name ="Text84"
                    ControlSource ="=IIf([DateLastSold]>#1/1/2000#,[DateLastSold],\"\")"
                    Format ="Short Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =1020
                    Height =225
                    TabIndex =8
                    Name ="Text86"
                    ControlSource ="StockNumber"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =225
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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4920
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text60"
                    ControlSource ="=Sum([Onhand])"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =60
                    Width =900
                    Height =225
                    TabIndex =2
                    Name ="Text61"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =285
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
                    ControlSource ="=Sum([Price]*[Onhand])"
                    Format ="Standard"

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
                    ControlSource ="=(Sum([Price]*[Onhand])-Sum([Onhand]*[Cost]))/Sum([Price]*[Onhand])"
                    Format ="Percent"

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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4020
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =285
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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =4920
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text68"
                    ControlSource ="=Sum([Onhand])"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =60
                    Width =900
                    Height =225
                    TabIndex =2
                    Name ="Text69"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =285
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
                    ControlSource ="=Sum([Price]*[Onhand])"
                    Format ="Standard"

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
                    ControlSource ="=(Sum([Price]*[Onhand])-Sum([Onhand]*[Cost]))/Sum([Price]*[Onhand])"
                    Format ="Percent"

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

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4020
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =285
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
                    Left =4920
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text26"
                    ControlSource ="=Sum([Onhand])"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =120
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="txtDepartmentCost"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =345
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
                    ControlSource ="=Sum([Price]*[Onhand])"
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
                    ControlSource ="=(Sum([Price]*[Onhand])-Sum([Onhand]*[Cost]))/Sum([Price]*[Onhand])"
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
                    Left =4020
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =120
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =345
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
                    Left =4920
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text36"
                    ControlSource ="=Sum([Onhand])"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =120
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =120
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =345
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
                    ControlSource ="=Sum([Price]*[Onhand])"
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
                    ControlSource ="=(Sum([Price]*[Onhand])-Sum([Onhand]*[Cost]))/Sum([Price]*[Onhand])"
                    Format ="Percent"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4020
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text82"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =120
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =345
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
                    Left =4740
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =345
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
                    ControlSource ="=Sum([Price]*[Onhand])"
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
                    ControlSource ="=(Sum([Price]*[Onhand])-Sum([Onhand]*[Cost]))/Sum([Price]*[Onhand])"
                    Format ="Percent"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3900
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text83"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =285
                End
            End
        End
    End
End
CodeBehindForm
' See "74 Items with Alpha UPC.cls"
