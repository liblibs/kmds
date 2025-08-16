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
    ItemSuffix =86
    Left =660
    Top =1410
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xb4c7bcc45bdee240
    End
    RecordSource ="qryItemOnhandChanges"
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
            ControlSource ="SalesDepartment.ID"
        End
        Begin BreakLevel
            ControlSource ="Item.ID"
        End
        Begin BreakLevel
            ControlSource ="SetupDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10680
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Audit Onhand Change"
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
                    Top =660
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4500
                    Top =960
                    Width =1920
                    Height =300
                    ColumnOrder =0
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text73"
                    ControlSource ="=Forms!PrintInventoryAuditReport!dtpDateEnd"
                    Format ="Medium Date"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =3840
                            Top =960
                            Width =525
                            Height =300
                            Name ="Label78"
                            Caption =" thru"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4500
                    Top =720
                    Width =1920
                    Height =300
                    ColumnOrder =1
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text76"
                    ControlSource ="=Forms!PrintInventoryAuditReport!dtpDateBegin"
                    Format ="Medium Date"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =2460
                            Top =660
                            Width =1920
                            Height =300
                            Name ="Label77"
                            Caption ="Range of dates from"
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =375
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =480
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =480
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =2640
                    Width =1530
                    Height =300
                    Name ="D2_Label"
                    Caption ="Adjustment Qty"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2640
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =300
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =30
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =330
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =360
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4290
                    Width =555
                    Height =300
                    Name ="Label57"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4290
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =4830
                    Width =900
                    Height =300
                    Name ="Label68"
                    Caption ="Ext Price"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4830
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =6660
                    Width =1200
                    Height =300
                    Name ="Label79"
                    Caption ="Changed by"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =9525
                    Width =1215
                    Height =300
                    Name ="Label80"
                    Caption ="Changed on"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =7860
                    Width =1110
                    Height =300
                    Name ="Label81"
                    Caption ="Reason"
                    LayoutCachedLeft =7860
                    LayoutCachedWidth =8970
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =5760
                    Width =840
                    Height =300
                    Name ="Label85"
                    Caption ="OnShelf"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5760
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =300
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =510
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Top =120
                    Width =588
                    Height =330
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Group"
                    ControlSource ="SalesDepartment.ID"
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
                    ControlSource ="SalesDepartment.Description"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Visible = NotDefault
                    Left =3900
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
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Width =960
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="Item.ID"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =960
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1020
                    Width =2760
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Item.Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =3810
                    Width =360
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Quantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3810
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4410
                    Width =540
                    Height =225
                    TabIndex =3
                    Name ="Text58"
                    ControlSource ="UPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4410
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5010
                    Width =660
                    Height =225
                    TabIndex =4
                    Name ="Ext"
                    ControlSource ="Ext"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5010
                    LayoutCachedWidth =5670
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6360
                    Width =420
                    Height =225
                    TabIndex =5
                    Name ="Cashier.ID"
                    ControlSource ="Cashier.ID"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Cashier_ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6780
                    Width =1140
                    Height =225
                    TabIndex =6
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8820
                    Width =1920
                    Height =225
                    ColumnWidth =2100
                    TabIndex =7
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="General Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7920
                    Width =1080
                    TabIndex =8
                    Name ="InputCode"
                    ControlSource ="InputCode"
                    FontName ="Arial Narrow"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000002000000000000001000000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b005400720061006e00730054007900700065005d003d002200520022000000 ,
                        0x0000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000020000000100000000000000ffff00000f0000005b00 ,
                        0x5400720061006e00730054007900700065005d003d0022005200220000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8700
                    Width =300
                    Height =225
                    TabIndex =9
                    Name ="TransType"
                    ControlSource ="TransType"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =5820
                    Width =480
                    Height =225
                    TabIndex =10
                    Name ="Text84"
                    ControlSource ="NewOnShelf"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5820
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =450
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =3420
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text66"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5160
                    Top =120
                    Width =840
                    FontWeight =700
                    TabIndex =1
                    Name ="Text67"
                    ControlSource ="=Sum([Ext])"
                    Format ="Standard"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1680
                    Top =120
                    Width =1503
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BackColor =12632256
                    Name ="Text71"
                    ControlSource ="SalesDepartment.Description"
                    FontName ="Times New Roman"

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
            Height =300
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =3420
                    Top =60
                    Width =720
                    Height =225
                    Name ="Text69"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =1080
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([Ext])"
                    Format ="Standard"

                End
                Begin Label
                    TextAlign =3
                    Width =1320
                    Height =300
                    Name ="Label72"
                    Caption ="Report total"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemOnhandChangeAudit.cls"
