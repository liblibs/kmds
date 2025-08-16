Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10380
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =585
    Top =1320
    RecSrcDt = Begin
        0x2546f79bd36de340
    End
    RecordSource ="qryPriceGroup"
    Caption ="Club"
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
            KeepTogether =1
            ControlSource ="PriceGroup.Description"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =480
                    Top =60
                    Width =2940
                    Height =510
                    FontSize =20
                    Name ="Label24"
                    Caption ="Price Groups"
                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =570
                End
            End
        End
        Begin PageHeader
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =360
                    Width =10260
                    Name ="Line27"
                    LayoutCachedTop =360
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =0
                    Left =4920
                    Width =630
                    Height =300
                    Name ="Label40"
                    Caption ="Price"
                    LayoutCachedLeft =4920
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =300
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =420
            Name ="GroupHeader0"
            AutoHeight =255
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Height =225
                    Name ="PriceGroup.Description"
                    ControlSource ="PriceGroup.Description"
                    EventProcPrefix ="PriceGroup_Description"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =285
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10200
                    Name ="Line5"
                    LayoutCachedTop =240
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2040
                    Width =2280
                    Height =225
                    Name ="Item.Description"
                    ControlSource ="Item.Description"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Item_Description"

                    LayoutCachedLeft =2040
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4680
                    Width =840
                    Height =225
                    TabIndex =1
                    Name ="Price"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4680
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Width =840
                    Height =225
                    TabIndex =2
                    Name ="Text41"
                    ControlSource ="Cost"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7440
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="Text42"
                    ControlSource ="Deposit"
                    Format ="#.00\\,000\\.00;-#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7440
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="Text43"
                    ControlSource ="QtyPack"
                    Format ="#,##0;-#,##0[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8400
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="Text44"
                    ControlSource ="QtyPackPrice"
                    Format ="#.00\\,000\\.00;-#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9360
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =420
                    Width =1320
                    Height =225
                    TabIndex =6
                    Name ="Text47"
                    ControlSource ="Item.ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =420
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Width =600
                    Height =225
                    TabIndex =7
                    Name ="Text50"
                    ControlSource ="GP"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            Name ="GroupFooter0"
            AutoHeight =255
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =3240
                    Top =60
                    Width =2820
                    Height =300
                    Name ="Text46"
                    ControlSource ="=\"Count =\" & Count([Item.Description])"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =360
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
                    Top =360
                    Width =4560
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =360
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =4440
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
                    Top =60
                    Width =4560
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =360
                    Width =10215
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedTop =360
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =180
                    Top =60
                    Width =10140
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =180
                    Top =120
                    Width =2820
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text73"
                    ControlSource ="=\"Count =\" & Count([Item.Description])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =420
                End
                Begin Line
                    Top =60
                    Width =10380
                    Name ="Line30"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =60
                End
                Begin Line
                    Top =120
                    Width =10380
                    Name ="Line31"
                    LayoutCachedTop =120
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =120
                End
            End
        End
    End
End
CodeBehindForm
' See "62 Price Group List.cls"
