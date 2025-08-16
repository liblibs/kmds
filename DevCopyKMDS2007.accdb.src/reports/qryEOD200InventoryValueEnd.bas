Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10756
    DatasheetFontHeight =10
    ItemSuffix =51
    Top =1245
    RecSrcDt = Begin
        0x200aa1452b8ae240
    End
    RecordSource ="qry EODInventoryValue"
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
            Height =1170
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10680
                    Height =540
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10680
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Value Report (End)"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =600
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =240
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3990
                    Top =120
                    Width =810
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    Left =5100
                    Top =120
                    Width =525
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5790
                    Top =120
                    Width =855
                    Height =300
                    Name ="D2_Label"
                    Caption ="Ext Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    Left =7290
                    Top =120
                    Width =615
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8130
                    Top =120
                    Width =1035
                    Height =300
                    Name ="D4_Label"
                    Caption =" Ext Retail"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10695
                    Height =30
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10695
                    Height =30
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =450
                    Width =10695
                    Height =30
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =480
                    Width =10695
                    Height =30
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =510
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =120
                    Top =120
                    Width =4983
                    Height =330
                    ColumnWidth =1185
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =450
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =300
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =60
                    Width =774
                    Height =300
                    ColumnWidth =1260
                    FontSize =11
                    ForeColor =8388608
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =834
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =900
                    Width =2823
                    Height =300
                    ColumnWidth =1710
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =3723
                    LayoutCachedHeight =300
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =255
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =180
                    Width =840
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =1140
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

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

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    Left =10140
                    Width =600
                    Height =255
                    ColumnWidth =840
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =4980
                    Width =1620
                    Height =225
                    ColumnWidth =735
                    TabIndex =4
                    Name ="ExtCost"
                    ControlSource ="=[Onhand]*[Cost]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =4800
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =5
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =9180
                    Width =300
                    Height =225
                    TabIndex =6
                    Name ="SalePriceItem"
                    ControlSource ="=IIf(IsNull([SalePrice]),\"\",\"*\")"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Visible = NotDefault
                    Left =9540
                    Width =660
                    Height =225
                    TabIndex =7
                    Name ="SalePrice"
                    ControlSource ="SalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =6660
                    Width =1200
                    Height =225
                    TabIndex =8
                    Name ="CurPrice"
                    ControlSource ="=IIf(IsNull([SalePrice]),[Price],[SalePrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =7620
                    Width =1620
                    Height =225
                    TabIndex =9
                    Name ="ExtCurPrice"
                    ControlSource ="=[CurPrice]*[Onhand]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =225
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4020
                    Width =900
                    Height =225
                    Name ="Text26"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =4980
                    Width =1620
                    Height =225
                    TabIndex =1
                    Name ="Text28"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =7620
                    Width =1620
                    Height =225
                    TabIndex =2
                    Name ="Text33"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =180
                    Width =780
                    Height =225
                    TabIndex =3
                    Name ="Text34"
                    ControlSource ="Department"

                    LayoutCachedLeft =180
                    LayoutCachedWidth =960
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =1020
                    Width =2940
                    Height =225
                    TabIndex =4
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4020
                    Top =60
                    Width =900
                    Height =225
                    Name ="Text36"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =4980
                    Width =1620
                    Height =225
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =7620
                    Top =60
                    Width =1620
                    Height =225
                    TabIndex =2
                    Name ="Text38"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =960
                    Top =60
                    Width =2880
                    Height =225
                    TabIndex =3
                    Name ="Text47"
                    ControlSource ="GroupDesc"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin PageFooter
            Height =630
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =5040
                    Height =300
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
                    Width =4980
                    Height =330
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
                    Height =300
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
                    Width =4980
                    Height =330
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
                    Left =60
                    Top =300
                    Width =10620
                    Height =330
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
                    Width =10620
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
            CanGrow = NotDefault
            Height =1380
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    Left =3240
                    Width =1620
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3240
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    Left =4980
                    Width =1620
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    Left =7620
                    Top =60
                    Width =1620
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text43"
                    ControlSource ="=Sum(IIf(IsNull([SalePrice]),[Price],[SalePrice])*[Onhand])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =345
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
                Begin Subform
                    Locked = NotDefault
                    Top =420
                    Width =10756
                    Height =960
                    TabIndex =3
                    Name ="qryEOS040TodaysLiquorRecap"
                    SourceObject ="Report.qryEOS040TodaysLiquorRecap"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD200InventoryValueEnd.cls"
