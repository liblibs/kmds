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
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =149
    Left =1380
    Top =270
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xdad125e1f1c8e540
    End
    RecordSource ="qryEO129RInvoicesWithCouponsLines"
    Caption ="EOD129RInvoicesWithCoupons"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            ControlSource ="InvoHeadID"
        End
        Begin BreakLevel
            ControlSource ="TransType"
        End
        Begin BreakLevel
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1380
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =90
                    Top =120
                    Width =10605
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Invoices with Coupons"
                    LayoutCachedLeft =90
                    LayoutCachedTop =120
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =660
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10680
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10680
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2700
                    Top =780
                    Height =315
                    ColumnOrder =2
                    FontSize =12
                    Name ="CouponItemID"
                    ControlSource ="CouponItemID"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =780
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1095
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Top =780
                    Width =2280
                    Height =315
                    ColumnOrder =1
                    FontSize =12
                    TabIndex =1
                    Name ="CouponDescription"
                    ControlSource ="CouponDescription"

                    LayoutCachedLeft =4260
                    LayoutCachedTop =780
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1095
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6600
                    Top =780
                    Height =315
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="Text139"
                    ControlSource ="=Sum([CouponTotal])"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =780
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1095
                End
            End
        End
        Begin PageHeader
            Height =660
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =120
                    Width =9600
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =60
                    Width =9600
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =9600
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =540
                    Width =9600
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =300
                    Top =180
                    Width =585
                    Height =300
                    FontWeight =400
                    Name ="Label128"
                    Caption ="Date "
                    LayoutCachedLeft =300
                    LayoutCachedTop =180
                    LayoutCachedWidth =885
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =0
                    Left =2400
                    Top =180
                    Width =780
                    Height =300
                    Name ="Label131"
                    Caption ="Receipt"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =180
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =0
                    Left =3780
                    Top =180
                    Width =855
                    Height =300
                    Name ="Label132"
                    Caption ="Register"
                    LayoutCachedLeft =3780
                    LayoutCachedTop =180
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =480
                End
                Begin Label
                    Left =5040
                    Top =180
                    Width =810
                    Height =300
                    Name ="Label133"
                    Caption ="Cashier"
                    LayoutCachedLeft =5040
                    LayoutCachedTop =180
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =480
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2400
                    Top =180
                    Height =225
                    FontWeight =700
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =180
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4035
                    Top =180
                    Width =600
                    Height =225
                    FontWeight =700
                    TabIndex =1
                    Name ="CashSysID"
                    ControlSource ="CashSysID"

                    LayoutCachedLeft =4035
                    LayoutCachedTop =180
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =180
                    Top =180
                    Width =2100
                    Height =225
                    ColumnWidth =2340
                    FontWeight =700
                    TabIndex =2
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    ShowDatePicker =1

                    LayoutCachedLeft =180
                    LayoutCachedTop =180
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5040
                    Top =180
                    Width =540
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="CashierID"
                    ControlSource ="CashierID"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =180
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5700
                    Top =180
                    Height =225
                    FontWeight =700
                    TabIndex =4
                    Name ="Name"
                    ControlSource ="Name"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =180
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =405
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =9600
                    Name ="Line127"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =9600
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7560
                    Top =180
                    Height =225
                    TabIndex =5
                    Name ="CouponTotal"
                    ControlSource ="CouponTotal"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =405
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1380
                    Height =225
                    Name ="Code"
                    ControlSource ="Code"
                    ConditionalFormat = Begin
                        0x0100000080000000010000000000000002000000000000000f00000001010001 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0043006f00750070006f006e004900740065006d00490044005d0000000000
                    End

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =225
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000101000100000000ffffff000e0000005b00 ,
                        0x43006f00750070006f006e004900740065006d00490044005d00000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Height =225
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6420
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Quantity"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Width =660
                    Height =225
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4440
                    Height =225
                    ColumnWidth =2700
                    TabIndex =4
                    Name ="InputCode"
                    ControlSource ="InputCode"

                    LayoutCachedLeft =4440
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5940
                    Width =360
                    Height =225
                    TabIndex =5
                    Name ="TransType"
                    ControlSource ="TransType"

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8280
                    Width =900
                    Height =225
                    TabIndex =6
                    Name ="IsCoupon"
                    ControlSource ="IsCoupon"

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =225
                End
            End
        End
        Begin PageFooter
            Height =690
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
                Begin Line
                    BorderWidth =3
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
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
                    Top =360
                    Width =10560
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

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =360
                    Width =10680
                    BorderColor =12632256
                    Name ="Line78"
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD129RInvoicesWithCoupons.cls"
