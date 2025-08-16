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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =158
    Left =1050
    Top =1860
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x209b1f711009e340
    End
    RecordSource ="qryTransferByDateByInvoHeadID"
    Caption ="Store Transfers"
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
            ControlSource ="InOrOut"
        End
        Begin BreakLevel
            ControlSource ="FirstOfSetupDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15
                    Top =60
                    Width =10755
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Store Transfers"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10785
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10785
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10621
                    Height =480
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin PageHeader
            Height =945
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10740
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10740
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =900
                    Width =10740
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =930
                    Width =10740
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =240
                    Width =2340
                    Height =300
                    Name ="Label58"
                    Caption ="*****    Liquor   *****"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =540
                    Width =2340
                    Height =300
                    Name ="Label59"
                    Caption ="Quantity  Retail     Cost"
                End
                Begin Label
                    TextAlign =3
                    Left =1440
                    Top =540
                    Width =570
                    Height =300
                    Name ="Label61"
                    Caption ="Store"
                End
                Begin Label
                    Left =540
                    Top =540
                    Width =900
                    Height =315
                    Name ="Label62"
                    Caption ="Transfer"
                End
                Begin Label
                    TextAlign =3
                    Left =6285
                    Top =240
                    Width =2115
                    Height =300
                    Name ="Label90"
                    Caption ="*** Non Liquor   ***"
                End
                Begin Label
                    TextAlign =3
                    Left =6300
                    Top =540
                    Width =2100
                    Height =300
                    Name ="Label91"
                    Caption ="Quantity  Retail  Cost"
                End
                Begin Label
                    TextAlign =3
                    Left =8400
                    Top =240
                    Width =2295
                    Height =300
                    Name ="Label101"
                    Caption ="***       Total    ***"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Top =540
                    Width =2220
                    Height =300
                    Name ="Label102"
                    Caption ="Quantity  Retail     Cost"
                End
                Begin Rectangle
                    Left =3780
                    Top =240
                    Width =2400
                    Height =600
                    Name ="Box118"
                End
                Begin Rectangle
                    Left =6300
                    Top =240
                    Width =2160
                    Height =600
                    Name ="Box119"
                End
                Begin Rectangle
                    Left =8460
                    Top =240
                    Width =2280
                    Height =600
                    Name ="Box120"
                End
                Begin Label
                    TextAlign =0
                    Top =540
                    Width =495
                    Height =300
                    Name ="Label130"
                    Caption ="Date"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Width =480
                    Height =315
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text156"
                    ControlSource ="InOrOut"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Width =1440
                    Height =315
                    Name ="Label157"
                    Caption ="Transfers"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =660
                    Width =720
                    Height =225
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Width =840
                    Height =225
                    TabIndex =1
                    Name ="LiquorRetail"
                    ControlSource ="LiquorRetail"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3960
                    Width =480
                    Height =225
                    TabIndex =2
                    Name ="LiquorQuantity"
                    ControlSource ="LiquorQuantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5340
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="LiquorCost"
                    ControlSource ="LiquorCost"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1680
                    Width =1980
                    TabIndex =4
                    Name ="InputCode"
                    ControlSource ="InputCode"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="NonLiquorRetail"
                    ControlSource ="NonLiquorRetail"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6240
                    Width =480
                    Height =225
                    TabIndex =6
                    Name ="NonLiquorQuantity"
                    ControlSource ="NonLiquorQuantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Width =840
                    Height =225
                    TabIndex =7
                    Name ="NonLiquorCost"
                    ControlSource ="NonLiquorCost"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Width =840
                    Height =225
                    TabIndex =8
                    Name ="Text103"
                    ControlSource ="TotalRetail"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Width =480
                    Height =225
                    TabIndex =9
                    Name ="Text104"
                    ControlSource ="TotalQuanity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Width =840
                    Height =225
                    TabIndex =10
                    Name ="Text105"
                    ControlSource ="TotalCost"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Width =600
                    Height =225
                    ColumnWidth =2400
                    TabIndex =11
                    Name ="SetupDate"
                    ControlSource ="FirstOfSetupDate"
                    Format ="mm/dd"

                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10800
                    Name ="Line5"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3660
                    Width =360
                    Height =225
                    TabIndex =12
                    Name ="InOrOut"
                    ControlSource ="InOrOut"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =480
                    Height =315
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text131"
                    ControlSource ="InOrOut"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =0
                    Left =2460
                    Top =60
                    Width =1440
                    Height =315
                    Name ="Label132"
                    Caption ="Transfer Total"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =1
                    Name ="Text133"
                    ControlSource ="=Sum([LiquorRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3960
                    Top =120
                    Width =480
                    Height =225
                    TabIndex =2
                    Name ="Text134"
                    ControlSource ="=Sum([LiquorQuantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5340
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="Text135"
                    ControlSource ="=Sum([LiquorCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="Text136"
                    ControlSource ="=Sum([NonLiquorRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6240
                    Top =120
                    Width =480
                    Height =225
                    TabIndex =5
                    Name ="Text137"
                    ControlSource ="=Sum([NonLiquorQuantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =6
                    Name ="Text138"
                    ControlSource ="=Sum([NonLiquorCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =7
                    Name ="Text139"
                    ControlSource ="=Sum([TotalRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =120
                    Width =480
                    Height =225
                    TabIndex =8
                    Name ="Text140"
                    ControlSource ="=Sum([TotalQuanity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =120
                    Width =840
                    Height =225
                    TabIndex =9
                    Name ="Text141"
                    ControlSource ="=Sum([TotalCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =3540
                    Top =60
                    Width =7200
                    Name ="Line154"
                End
                Begin Line
                    Left =3540
                    Top =420
                    Width =7200
                    Name ="Line155"
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =10680
                    Height =300
                    FontWeight =700
                    TabIndex =3
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
                    TabIndex =4
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5760
                    Width =4965
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =540
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =1800
                    Top =240
                    Width =1080
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin Line
                    Left =3540
                    Top =60
                    Width =7200
                    Name ="Line70"
                End
                Begin Label
                    Visible = NotDefault
                    Top =240
                    Width =2100
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
                Begin Line
                    Left =3540
                    Top =120
                    Width =7200
                    Name ="Line81"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Top =240
                    Width =840
                    Height =225
                    Name ="Text145"
                    ControlSource ="=Sum([LiquorRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3960
                    Top =240
                    Width =480
                    Height =225
                    TabIndex =1
                    Name ="Text146"
                    ControlSource ="=Sum([LiquorQuantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5340
                    Top =240
                    Width =840
                    Height =225
                    TabIndex =2
                    Name ="Text147"
                    ControlSource ="=Sum([LiquorCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Top =240
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="Text148"
                    ControlSource ="=Sum([NonLiquorRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6240
                    Top =240
                    Width =480
                    Height =225
                    TabIndex =4
                    Name ="Text149"
                    ControlSource ="=Sum([NonLiquorQuantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Top =240
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="Text150"
                    ControlSource ="=Sum([NonLiquorCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Top =240
                    Width =840
                    Height =225
                    TabIndex =6
                    Name ="Text151"
                    ControlSource ="=Sum([TotalRetail])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =240
                    Width =480
                    Height =225
                    TabIndex =7
                    Name ="Text152"
                    ControlSource ="=Sum([TotalQuanity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =240
                    Width =840
                    Height =225
                    TabIndex =8
                    Name ="Text153"
                    ControlSource ="=Sum([TotalCost])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "44 Store Transfers by Invoice.cls"
