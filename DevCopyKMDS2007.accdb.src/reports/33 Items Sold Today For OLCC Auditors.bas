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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =102
    Top =60
    OnNoData ="[Event Procedure]"
    OrderBy ="ID"
    RecSrcDt = Begin
        0x5c0415afef1ce440
    End
    RecordSource ="qryItemsSoldToday"
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
            Height =1110
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    Left =60
                    Top =600
                    Width =4861
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =5040
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items sold today"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Width =10560
                    Height =60
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1080
                    Width =10560
                    Height =30
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
                    Width =1125
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sold Today"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =120
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =60
                    Top =60
                    Width =10500
                    Height =30
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6705
                    Top =120
                    Width =495
                    Height =300
                    Name ="Label94"
                    Caption ="CR1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6705
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =7545
                    Top =120
                    Width =495
                    Height =300
                    Name ="Label95"
                    Caption ="CR2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7545
                    LayoutCachedTop =120
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =8385
                    Top =120
                    Width =495
                    Height =300
                    Name ="Label96"
                    Caption ="CR3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8385
                    LayoutCachedTop =120
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =9165
                    Top =120
                    Width =495
                    Height =300
                    Name ="Label97"
                    Caption ="CR4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9165
                    LayoutCachedTop =120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =2
                    Left =120
                    Top =480
                    Width =10500
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9945
                    Top =120
                    Width =495
                    Height =300
                    Name ="Label98"
                    Caption ="CR5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9945
                    LayoutCachedTop =120
                    LayoutCachedWidth =10440
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
        Begin Section
            KeepTogether = NotDefault
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
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
                    Left =1440
                    Width =2760
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =1020
                    Top =240
                    Width =9060
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6525
                    Width =795
                    TabIndex =2
                    Name ="SoldTodayCr1"
                    ControlSource ="SoldTodayCr1"
                    Format ="#,##0;(#,##0)[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6525
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7305
                    Width =795
                    TabIndex =3
                    Name ="SoldTodayCr2"
                    ControlSource ="SoldTodayCr2"
                    Format ="#,##0;(#,##0)[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7305
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8085
                    Width =795
                    TabIndex =4
                    Name ="SoldTodayCr3"
                    ControlSource ="SoldTodayCr3"
                    Format ="#,##0;(#,##0)[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8085
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8865
                    Width =795
                    TabIndex =5
                    Name ="SoldTodayCr4"
                    ControlSource ="SoldTodayCr4"
                    Format ="#,##0;(#,##0)[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8865
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Width =780
                    TabIndex =6
                    Name ="SoldToday"
                    ControlSource ="SoldToday"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4500
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9645
                    Width =795
                    TabIndex =7
                    Name ="Text99"
                    ControlSource ="SoldTodayCr5"
                    Format ="#,##0;(#,##0)[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9645
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =720
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2040
                    Top =120
                    Width =420
                    Height =225
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
                    TabIndex =1
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6525
                    Top =120
                    Width =795
                    TabIndex =3
                    Name ="Text84"
                    ControlSource ="=Sum([SoldTodayCr1])"

                    LayoutCachedLeft =6525
                    LayoutCachedTop =120
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7305
                    Top =120
                    Width =795
                    TabIndex =4
                    Name ="Text85"
                    ControlSource ="=Sum([SoldTodayCr2])"

                    LayoutCachedLeft =7305
                    LayoutCachedTop =120
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8085
                    Top =120
                    Width =795
                    TabIndex =5
                    Name ="Text86"
                    ControlSource ="=Sum([SoldTodayCr3])"

                    LayoutCachedLeft =8085
                    LayoutCachedTop =120
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9090
                    Top =120
                    Width =795
                    TabIndex =6
                    Name ="Text87"
                    ControlSource ="=Sum([SoldTodayCr4])"

                    LayoutCachedLeft =9090
                    LayoutCachedTop =120
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =120
                    Width =1020
                    TabIndex =7
                    Name ="Text88"
                    ControlSource ="=Sum([SoldToday])"

                    LayoutCachedLeft =4260
                    LayoutCachedTop =120
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9645
                    Top =120
                    Width =795
                    TabIndex =8
                    Name ="Text100"
                    ControlSource ="=Sum([SoldTodayCr5])"

                    LayoutCachedLeft =9645
                    LayoutCachedTop =120
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =360
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
                    Left =960
                    Top =120
                    Width =2880
                    Height =225
                    Name ="Text47"
                    ControlSource ="GroupDesc"

                End
            End
        End
        Begin PageFooter
            Height =615
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
                    Width =4845
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
                    Width =4845
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
                    Width =10425
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
            Height =660
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6300
                    Top =60
                    Width =1020
                    Name ="Text89"
                    ControlSource ="=Sum([SoldTodayCr1])"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =60
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =1020
                    TabIndex =1
                    Name ="Text90"
                    ControlSource ="=Sum([SoldTodayCr2])"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7860
                    Top =60
                    Width =1020
                    TabIndex =2
                    Name ="Text91"
                    ControlSource ="=Sum([SoldTodayCr3])"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =60
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8865
                    Top =60
                    Width =1020
                    TabIndex =3
                    Name ="Text92"
                    ControlSource ="=Sum([SoldTodayCr4])"

                    LayoutCachedLeft =8865
                    LayoutCachedTop =60
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =60
                    Width =1020
                    TabIndex =4
                    Name ="Text93"
                    ControlSource ="=Sum([SoldToday])"

                    LayoutCachedLeft =4260
                    LayoutCachedTop =60
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9780
                    Top =60
                    Width =660
                    TabIndex =5
                    Name ="Text101"
                    ControlSource ="=Sum([SoldTodayCr5])"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =300
                End
            End
        End
    End
End
CodeBehindForm
' See "33 Items Sold Today For OLCC Auditors.cls"
