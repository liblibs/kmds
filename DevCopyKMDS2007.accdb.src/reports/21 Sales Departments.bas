Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =1050
    Top =1860
    RecSrcDt = Begin
        0x58dd1178d16de340
    End
    RecordSource ="qrySalesDepartments"
    Caption ="qrySalesDepartments"
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
            ControlSource ="GroupLevel"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SalesGroup.ID"
        End
        Begin BreakLevel
            ControlSource ="SalesDepartment.ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10020
                    Height =510
                    FontSize =20
                    Name ="Label20"
                    Caption ="Sales Departments"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line23"
                End
            End
        End
        Begin PageHeader
            Height =1380
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10020
                    Name ="Line24"
                    LayoutCachedWidth =10020
                End
                Begin Line
                    BorderWidth =2
                    Top =720
                    Width =10020
                    Name ="Line25"
                    LayoutCachedTop =720
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =720
                End
                Begin Label
                    TextAlign =0
                    Left =3720
                    Top =360
                    Width =990
                    Height =300
                    Name ="Label29"
                    Caption ="Discounts"
                End
                Begin Label
                    TextAlign =2
                    Left =4733
                    Top =60
                    Width =1020
                    Height =600
                    Name ="Label30"
                    Caption ="Neg OH Report"
                End
                Begin Label
                    TextAlign =2
                    Left =5880
                    Top =60
                    Width =990
                    Height =540
                    Name ="Label32"
                    Caption =" Allow Discounts"
                End
                Begin Label
                    TextAlign =0
                    Left =6960
                    Top =60
                    Width =1140
                    Height =600
                    Name ="Label34"
                    Caption ="Inventory Value Rpt"
                End
                Begin Label
                    TextAlign =2
                    Left =8160
                    Top =60
                    Width =900
                    Height =600
                    Name ="Label36"
                    Caption ="Target GP"
                End
                Begin Label
                    TextAlign =2
                    Left =9180
                    Top =60
                    Width =900
                    Height =600
                    Name ="Label39"
                    Caption ="SKU Count"
                    LayoutCachedLeft =9180
                    LayoutCachedTop =60
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =660
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =390
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =2640
                    Height =300
                    ColumnWidth =2055
                    FontSize =11
                    ForeColor =8388608
                    Name ="Text26"
                    ControlSource ="SalesGroup.Description"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =360
                    Width =660
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="SalesDepartment.ID"
                    ControlSource ="SalesDepartment.ID"
                    EventProcPrefix ="SalesDepartment_ID"

                    LayoutCachedLeft =360
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1080
                    Width =2580
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="SalesDepartment.Description"
                    ControlSource ="SalesDepartment.Description"
                    EventProcPrefix ="SalesDepartment_Description"

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8280
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="TargetGP"
                    ControlSource ="TargetGP"
                    Format ="Percent"

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =225
                End
                Begin CheckBox
                    Left =3960
                    Top =60
                    TabIndex =3
                    Name ="Check28"
                    ControlSource ="Discounts"

                End
                Begin CheckBox
                    Left =5040
                    Top =60
                    TabIndex =4
                    Name ="Check31"
                    ControlSource ="NegativeOnhand"

                End
                Begin CheckBox
                    Left =6120
                    Top =60
                    TabIndex =5
                    Name ="Check33"
                    ControlSource ="AllowDiscounts"

                End
                Begin CheckBox
                    Left =7200
                    Top =60
                    TabIndex =6
                    Name ="Check35"
                    ControlSource ="IncludeInInventoryValueReports"

                End
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =9300
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Text40"
                    ControlSource ="Items"
                    Format ="Standard"

                    LayoutCachedLeft =9300
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =195
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin Line
                    Left =420
                    Top =60
                    Width =9600
                    Name ="Line37"
                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =60
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
                    Width =5340
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
