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
    ItemSuffix =87
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xaf7bb57fd6f3e540
    End
    RecordSource ="qryEOTenderByDepartmentByRegister"
    Caption ="qryEOD040SalesByDepartment1"
    DatasheetFontName ="Arial"
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
            ControlSource ="Department"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Department"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
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
                    Left =2145
                    Top =60
                    Width =6480
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Tender Totals by Register Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10800
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10800
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =495
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =1080
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3120
                    Top =120
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3120
                    LayoutCachedTop =120
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =4500
                    Top =120
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =120
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =5340
                    Top =120
                    Width =1020
                    Height =300
                    Name ="D1_Label"
                    Caption =" Reg 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5340
                    LayoutCachedTop =120
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =6420
                    Top =120
                    Width =1020
                    Height =300
                    Name ="D2_Label"
                    Caption ="Reg 2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6420
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =7500
                    Top =120
                    Width =1020
                    Height =300
                    Name ="D3_Label"
                    Caption ="Reg 3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7500
                    LayoutCachedTop =120
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =8640
                    Top =120
                    Width =1020
                    Height =300
                    Name ="D4_Label"
                    Caption =" Reg 4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8640
                    LayoutCachedTop =120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =150
                    Width =9720
                    Name ="Line54"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =150
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =150
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =120
                    Width =9720
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =450
                    Width =9720
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =9720
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9660
                    Top =120
                    Width =1020
                    Height =300
                    Name ="Label86"
                    Caption =" Reg 5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =420
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader0"
            AutoHeight =255
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =3
                    Left =120
                    Width =840
                    Height =270
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =960
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =3
                    Left =1020
                    Width =1740
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    Left =2880
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="TotalQ"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4080
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4080
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =5340
                    Width =1020
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="D1"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5340
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6420
                    Width =1020
                    Height =270
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="D2"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7500
                    Width =1020
                    Height =270
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="D3"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8640
                    Width =1020
                    Height =270
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="D4"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8640
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9660
                    Width =1020
                    Height =270
                    TabIndex =8
                    Name ="Text68"
                    ControlSource ="D5"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =270
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter1"
            AutoHeight =255
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2880
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="Text70"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4080
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text71"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4080
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5340
                    Width =1020
                    Height =300
                    TabIndex =2
                    Name ="Text72"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5340
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6420
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="Text73"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7500
                    Width =1020
                    Height =300
                    TabIndex =4
                    Name ="Text74"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8640
                    Width =1020
                    Height =300
                    TabIndex =5
                    Name ="Text75"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8640
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9660
                    Width =1020
                    Height =300
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Left =120
                    Width =840
                    Height =270
                    TabIndex =7
                    Name ="Text77"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =960
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Left =1020
                    Width =1740
                    Height =270
                    TabIndex =8
                    Name ="Text78"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =660
            Name ="PageFooterSection"
            Begin
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
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line60"
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
                    Left =120
                    Top =360
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
                    Left =120
                    Top =60
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
            Height =840
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextFontCharSet =204
                    TextFontFamily =34
                    Top =120
                    Width =1155
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Tender Total"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    Visible = NotDefault
                    Top =420
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
                Begin Line
                    Top =60
                    Width =10800
                    Name ="Line67"
                End
                Begin TextBox
                    OverlapFlags =4
                    TextFontCharSet =204
                    TextAlign =3
                    Left =2880
                    Top =120
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Text79"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =4080
                    Top =120
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text80"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =120
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =5340
                    Top =120
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text81"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =120
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =6420
                    Top =120
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text82"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =120
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =7500
                    Top =120
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text83"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =120
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =8640
                    Top =120
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="Text84"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =204
                    TextAlign =3
                    Left =9660
                    Top =120
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="Text85"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD040SalesByTenderByRegister.cls"
