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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =80
    Left =7140
    Top =408
    RecSrcDt = Begin
        0xd8c6d9d114a0e440
    End
    RecordSource ="qryDepositByItem"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            ControlSource ="Item.Description"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
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
                    Width =7350
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items with Deposit Inventory Value Report"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7410
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
                    Left =4560
                    Top =120
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4560
                    LayoutCachedTop =120
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =5670
                    Top =120
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5670
                    LayoutCachedTop =120
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =6360
                    Top =120
                    Width =855
                    Height =300
                    Name ="D2_Label"
                    Caption ="Ext Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6360
                    LayoutCachedTop =120
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =420
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
            Height =375
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =672
                    Height =375
                    ColumnWidth =1260
                    FontSize =14
                    Name ="Department"
                    ControlSource ="SalesDepartmentID"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =732
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =780
                    Width =2304
                    Height =375
                    ColumnWidth =1710
                    FontSize =14
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="SalesDepartment.Description"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =780
                    LayoutCachedWidth =3084
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =120
                    Width =1380
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1620
                    Width =2760
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Item.Description"
                    FontName ="Arial Narrow"
                    ConditionalFormat = Begin
                        0x01000000aa000000010000000100000000000000000000002400000001010100 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530075006200430061007400650067006f00720079004400650073006300 ,
                        0x720069007000740069006f006e005d003d0022004f00420053004f004c004500 ,
                        0x54004500220000000000
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101010000000000ffffff00230000005b00 ,
                        0x530075006200430061007400650067006f007200790044006500730063007200 ,
                        0x69007000740069006f006e005d003d0022004f00420053004f004c0045005400 ,
                        0x45002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    Left =4740
                    Width =600
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    Format ="#,##0;-#,##0[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4740
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6600
                    Width =900
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="ExtCost"
                    ControlSource ="DepositTotal"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =5640
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Deposit"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =240
                    Top =240
                    Width =10380
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
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
                    ControlSource ="=Sum([DepositTotal])"
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
                    ControlSource ="SalesDepartmentID"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    Left =60
                    Top =120
                    Width =1920
                    Height =225
                    TabIndex =3
                    Name ="Text35"
                    ControlSource ="SalesDepartment.Description"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3060
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"

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
                    ControlSource ="=Sum([DepositTotal])"
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
            End
        End
    End
End
CodeBehindForm
' See "73 Bottle Deposit Inventory Value.cls"
