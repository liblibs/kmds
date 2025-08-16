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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =55
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x919537505aa6e240
    End
    RecordSource ="qryEOItemsSoldToday"
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
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10695
                    Height =510
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
                    Left =3945
                    Top =60
                    Width =2955
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items Sold Today"
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
                    Left =4320
                    Top =120
                    Width =480
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5895
                    Top =120
                    Width =750
                    Height =300
                    Name ="D2_Label"
                    Caption ="Dollars"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
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
            Height =510
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Top =120
                    Width =588
                    Height =330
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Group"
                    ControlSource ="Group"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =720
                    Top =120
                    Width =1503
                    Height =330
                    ColumnWidth =1185
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Visible = NotDefault
                    Left =2400
                    Top =120
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
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
                    TextFontFamily =18
                    Left =180
                    Width =504
                    Height =300
                    ColumnWidth =1260
                    FontSize =11
                    ForeColor =8388608
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =660
                    Width =1728
                    Height =300
                    ColumnWidth =1710
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =180
                    Width =840
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    Left =1140
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    Left =4020
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="SoldToday"
                    ControlSource ="SoldToday"

                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =5700
                    Width =900
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="DollarsSoldToday"
                    ControlSource ="DollarsSoldToday"
                    Format ="Standard"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =4740
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text26"
                    ControlSource ="=Sum([SoldToday])"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =6420
                    Top =120
                    Width =960
                    Height =225
                    FontWeight =700
                    TabIndex =1
                    Name ="Text28"
                    ControlSource ="=Sum([DollarsSoldToday])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =900
                    Top =120
                    Width =420
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="Department"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =1380
                    Top =120
                    Width =2940
                    Height =225
                    TabIndex =3
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"

                End
                Begin Label
                    Left =7440
                    Top =120
                    Width =180
                    Height =300
                    Name ="Label51"
                    Caption ="*"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5640
                    Top =180
                    Width =720
                    Height =225
                    Name ="Text36"
                    ControlSource ="=Sum([SoldToday])"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =7140
                    Top =180
                    Width =1080
                    Height =225
                    FontWeight =700
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([DollarsSoldToday])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =3300
                    Top =180
                    Width =2160
                    Height =225
                    TabIndex =2
                    Name ="Text47"
                    ControlSource ="GroupDesc"

                End
                Begin Label
                    Left =8280
                    Top =180
                    Width =300
                    Height =300
                    Name ="Label52"
                    Caption ="**"
                End
            End
        End
        Begin PageFooter
            Height =600
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
                    Width =5040
                    Height =300
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
                    Width =5040
                    Height =300
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
                    Width =10680
                    Height =300
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
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =6480
                    Top =60
                    Width =720
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([SoldToday])"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =7740
                    Top =60
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([DollarsSoldToday])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Left =3060
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin Label
                    Left =9120
                    Top =60
                    Width =405
                    Height =300
                    Name ="Label54"
                    Caption ="***"
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOX120RItemsSoldToday.cls"
