Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
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
    ItemSuffix =144
    Left =1320
    Top =285
    OnNoData ="[Event Procedure]"
    OrderBy ="ID"
    RecSrcDt = Begin
        0xb438d9929d4fe640
    End
    RecordSource ="qry79ItemRestockListByCategory"
    Caption ="Items GP Report"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
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
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Line
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            OldBorderStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="SalesDptDescr"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Category"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            CanGrow = NotDefault
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =60
                    Width =10800
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="End of Day Item Restock List By Category"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =600
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line51"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =60
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10800
                    BorderColor =12632256
                    Name ="Line53"
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =960
                    Width =10800
                    Height =300
                    ColumnOrder =1
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text125"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                    LayoutCachedTop =960
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =660
                    Width =10800
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text126"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                    LayoutCachedTop =660
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =960
                End
            End
        End
        Begin PageHeader
            Height =1500
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =60
                    Width =10800
                    Height =540
                    FontSize =20
                    Name ="Label139"
                    Caption ="End of Day Item Restock List By Category"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =600
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line140"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =60
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10800
                    BorderColor =12632256
                    Name ="Line141"
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =960
                    Width =10800
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="lblTime"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                    LayoutCachedTop =960
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =660
                    Width =10800
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text143"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                    LayoutCachedTop =660
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =960
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =1800
            Name ="GroupHeader1"
            Begin
                Begin Label
                    TextAlign =0
                    Left =8460
                    Top =900
                    Width =2295
                    Height =659
                    TopMargin =158
                    Name ="Label128"
                    Caption ="Discrepency/Comments"
                    LayoutCachedLeft =8460
                    LayoutCachedTop =900
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =1559
                End
                Begin Label
                    TextAlign =0
                    Left =120
                    Top =900
                    Width =1194
                    Height =659
                    TopMargin =158
                    Name ="Department_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =900
                    LayoutCachedWidth =1314
                    LayoutCachedHeight =1559
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =900
                    Width =10800
                    Name ="Line54"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =900
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =900
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =840
                    Width =10800
                    Name ="Line55"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =840
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =840
                End
                Begin Line
                    BorderWidth =1
                    Top =1559
                    Width =10800
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =1559
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1559
                End
                Begin Line
                    BorderWidth =1
                    Top =1620
                    Width =10800
                    Name ="Line57"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =1620
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Top =240
                    Width =10800
                    Height =432
                    FontSize =15
                    FontWeight =700
                    BackColor =13952764
                    ForeColor =8388608
                    Name ="Text132"
                    ControlSource ="SalesDepartment"
                    FontName ="Times New Roman"

                    LayoutCachedTop =240
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =672
                End
                Begin Label
                    TextAlign =2
                    Left =5199
                    Top =900
                    Width =890
                    Height =659
                    Name ="Label114"
                    Caption ="Restock Amount"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5199
                    LayoutCachedTop =900
                    LayoutCachedWidth =6089
                    LayoutCachedHeight =1559
                End
                Begin Label
                    Left =2199
                    Top =900
                    Width =3000
                    Height =659
                    TopMargin =158
                    Name ="Label116"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2199
                    LayoutCachedTop =900
                    LayoutCachedWidth =5199
                    LayoutCachedHeight =1559
                End
                Begin Label
                    TextAlign =3
                    Left =1335
                    Top =900
                    Width =864
                    Height =659
                    TopMargin =158
                    RightMargin =72
                    Name ="Label120"
                    Caption ="Size"
                    LayoutCachedLeft =1335
                    LayoutCachedTop =900
                    LayoutCachedWidth =2199
                    LayoutCachedHeight =1559
                End
                Begin Label
                    TextAlign =2
                    Left =7011
                    Top =900
                    Width =1545
                    Height =659
                    TopMargin =158
                    Name ="Label121"
                    Caption ="Category"
                    LayoutCachedLeft =7011
                    LayoutCachedTop =900
                    LayoutCachedWidth =8556
                    LayoutCachedHeight =1559
                End
                Begin Label
                    TextAlign =3
                    Left =6089
                    Top =900
                    Width =864
                    Height =659
                    TopMargin =158
                    Name ="Label123"
                    Caption ="OnShelf"
                    LayoutCachedLeft =6089
                    LayoutCachedTop =900
                    LayoutCachedWidth =6953
                    LayoutCachedHeight =1559
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =435
            BreakLevel =1
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Top =60
                    Width =10800
                    Height =375
                    FontSize =14
                    FontWeight =700
                    LeftMargin =72
                    ForeColor =8388608
                    Name ="Location"
                    ControlSource ="Category"
                    FontName ="Times New Roman"

                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =435
                    BackThemeColorIndex =3
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =420
                    Width =10800
                    Name ="Line131"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =420
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =1194
                    Height =285
                    FontSize =10
                    Name ="Text82"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1314
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Left =2199
                    Top =60
                    Width =3000
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text83"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2199
                    LayoutCachedTop =60
                    LayoutCachedWidth =5199
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5199
                    Top =60
                    Width =890
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="RestockAmount"
                    ControlSource ="RestockAmount"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5199
                    LayoutCachedTop =60
                    LayoutCachedWidth =6089
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7011
                    Top =60
                    Width =1545
                    Height =285
                    TabIndex =3
                    Name ="SubCategory.Description"
                    ControlSource ="Category"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="SubCategory_Description"

                    LayoutCachedLeft =7011
                    LayoutCachedTop =60
                    LayoutCachedWidth =8556
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1335
                    Top =60
                    Width =864
                    Height =285
                    FontSize =10
                    TabIndex =4
                    RightMargin =72
                    Name ="Size"
                    ControlSource ="Size"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1335
                    LayoutCachedTop =60
                    LayoutCachedWidth =2199
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6089
                    Top =60
                    Width =864
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6089
                    LayoutCachedTop =60
                    LayoutCachedWidth =6953
                    LayoutCachedHeight =345
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =180
            BreakLevel =1
            Name ="GroupFooter1"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =180
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =1260
            Name ="PageFooterSection"
            Begin
                Begin Label
                    OldBorderStyle =1
                    Left =180
                    Top =780
                    Width =4725
                    Height =420
                    Name ="Label129"
                    Caption ="Completed by: ___________________________"
                    LayoutCachedLeft =180
                    LayoutCachedTop =780
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OldBorderStyle =1
                    Left =5460
                    Top =780
                    Width =5055
                    Height =420
                    Name ="Label130"
                    Caption ="Time Completed : ___________________________"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =780
                    LayoutCachedWidth =10515
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
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
                    Left =5760
                    Top =60
                    Width =4920
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
                    Left =120
                    Top =60
                    Width =10560
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =360
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
                    Left =5760
                    Top =360
                    Width =4920
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
                    Top =360
                    Width =10560
                    Height =330
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =690
                End
                Begin Line
                    BorderWidth =3
                    Top =15
                    Width =10800
                    BorderColor =12632256
                    Name ="Line133"
                    LayoutCachedTop =15
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =15
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =780
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextFontFamily =34
                    Top =120
                    Width =1620
                    Height =317
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Total"
                    FontName ="Arial Narrow"
                    LayoutCachedTop =120
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =437
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =960
                    Top =120
                    Width =1980
                    Height =317
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =960
                    LayoutCachedTop =120
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =437
                End
                Begin Line
                    BorderWidth =3
                    Top =480
                    Width =10800
                    BorderColor =12632256
                    Name ="Line78"
                    LayoutCachedTop =480
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =162
                    TextAlign =3
                    Left =4937
                    Top =120
                    Width =1152
                    Height =317
                    FontSize =10
                    FontWeight =700
                    Name ="Text110"
                    ControlSource ="=Sum([RestockAmount])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4937
                    LayoutCachedTop =120
                    LayoutCachedWidth =6089
                    LayoutCachedHeight =437
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =162
                    TextAlign =3
                    Left =7560
                    Top =120
                    Width =1152
                    Height =317
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text111"
                    ControlSource ="=Sum([Dollars])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =120
                    LayoutCachedWidth =8712
                    LayoutCachedHeight =437
                End
            End
        End
    End
End
CodeBehindForm
' See "79 Item Restock List By Category.cls"
