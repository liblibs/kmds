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
    ItemSuffix =83
    Left =180
    Top =615
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x03345fea4970e340
    End
    RecordSource ="qryItemsAfterRecalcOfUnitsPerOrder"
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
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10681
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =60
                    Width =10785
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items After Recalc of Units Per Order"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10785
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
            Height =855
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =240
                    Top =540
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =240
                    LayoutCachedTop =540
                    LayoutCachedWidth =1395
                    LayoutCachedHeight =840
                End
                Begin Label
                    Left =3990
                    Top =540
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3990
                    LayoutCachedTop =540
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =840
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
                    Top =840
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =840
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =840
                End
                Begin Label
                    TextAlign =2
                    Left =7260
                    Top =180
                    Width =1380
                    Height =660
                    Name ="D1_Label"
                    Caption ="Before Units Per Order"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7260
                    LayoutCachedTop =180
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =840
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =180
                    Width =1380
                    Height =660
                    Name ="Label79"
                    Caption ="After Units Per Order"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =180
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =840
                End
                Begin Label
                    TextAlign =2
                    Left =5760
                    Top =180
                    Width =1380
                    Height =660
                    Name ="Label81"
                    Caption ="Units Per Case"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =180
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =840
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =60
                    Width =1320
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1440
                    Width =2580
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =225
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
                    DecimalPlaces =0
                    TextAlign =3
                    Left =7680
                    Width =840
                    Height =225
                    ColumnWidth =735
                    TabIndex =3
                    Name ="Cost"
                    ControlSource ="AvgTurns"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7680
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =240
                    Width =9960
                    Name ="Line5"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =9120
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="Text80"
                    ControlSource ="UnitsPerOrder"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9120
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =6240
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="Text82"
                    ControlSource ="AvgTurns"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6240
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =225
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
                    Width =5040
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
' See "Items After Recalc of Units Per Order.cls"
