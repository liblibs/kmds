Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =10
    ItemSuffix =143
    Top =1245
    RecSrcDt = Begin
        0x08a4d751cdb4e540
    End
    RecordSource ="qryBottleHistoryByClubMonth"
    Caption ="qryClubHistoryByBottleByMonth"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =255
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
            ControlSource ="CatDesc"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Code"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1080
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =7320
                    Height =510
                    FontSize =20
                    Name ="Label70"
                    Caption ="Bottle History By Club By Month"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =14400
                    BorderColor =12632256
                    Name ="Line73"
                End
                Begin TextBox
                    TextAlign =2
                    Top =660
                    Width =14341
                    Height =300
                    ColumnOrder =0
                    Name ="Text97"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =510
                    Height =300
                    Name ="ClubID_Label"
                    Caption ="Club"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =780
                    Top =60
                    Width =720
                    Height =300
                    Name ="Code_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6465
                    Top =60
                    Width =600
                    Height =300
                    Name ="01_Label"
                    Caption ="01"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl01_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =7125
                    Top =60
                    Width =600
                    Height =300
                    Name ="02_Label"
                    Caption ="02"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl02_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =7785
                    Top =60
                    Width =600
                    Height =300
                    Name ="03_Label"
                    Caption ="03"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl03_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =8445
                    Top =60
                    Width =600
                    Height =300
                    Name ="04_Label"
                    Caption ="04"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl04_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9105
                    Top =60
                    Width =600
                    Height =300
                    Name ="05_Label"
                    Caption ="05"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl05_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9765
                    Top =60
                    Width =600
                    Height =300
                    Name ="06_Label"
                    Caption ="06"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl06_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =10425
                    Top =60
                    Width =600
                    Height =300
                    Name ="07_Label"
                    Caption ="07"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl07_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =11085
                    Top =60
                    Width =600
                    Height =300
                    Name ="08_Label"
                    Caption ="08"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl08_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =11745
                    Top =60
                    Width =600
                    Height =300
                    Name ="09_Label"
                    Caption ="09"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl09_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =12405
                    Top =60
                    Width =600
                    Height =300
                    Name ="10_Label"
                    Caption ="10"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl10_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =13065
                    Top =60
                    Width =600
                    Height =300
                    Name ="11_Label"
                    Caption ="11"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl11_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =13725
                    Top =60
                    Width =600
                    Height =300
                    Name ="12_Label"
                    Caption ="12"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl12_Label"
                End
                Begin Line
                    BorderWidth =2
                    Width =14400
                    Name ="Line74"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =14400
                    Name ="Line75"
                End
                Begin Label
                    TextAlign =3
                    Left =5520
                    Top =60
                    Width =885
                    Height =300
                    Name ="Label76"
                    Caption ="TotalQty"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =1680
                    Top =60
                    Width =1095
                    Height =300
                    Name ="Label93"
                    Caption ="Descripton"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3900
                    Top =60
                    Width =720
                    Height =300
                    Name ="Label94"
                    Caption ="Size"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4620
                    Top =60
                    Width =885
                    Height =300
                    Name ="Label98"
                    Caption ="Total $"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =345
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =5820
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="CatDesc"
                    ControlSource ="CatDesc"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =345
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =300
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =60
                    Width =720
                    Height =225
                    Name ="Code"
                    ControlSource ="Code"

                    LayoutCachedTop =60
                    LayoutCachedWidth =720
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =780
                    Top =60
                    Width =2400
                    Height =225
                    TabIndex =1
                    Name ="Text96"
                    ControlSource ="Description"

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =285
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Width =600
                    Height =225
                    Name ="01"
                    ControlSource ="01"
                    Format ="Standard"
                    EventProcPrefix ="Ctl01"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7125
                    Width =600
                    Height =225
                    TabIndex =1
                    Name ="02"
                    ControlSource ="02"
                    Format ="Standard"
                    EventProcPrefix ="Ctl02"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7785
                    Width =600
                    Height =225
                    TabIndex =2
                    Name ="03"
                    ControlSource ="03"
                    Format ="Standard"
                    EventProcPrefix ="Ctl03"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8445
                    Width =600
                    Height =225
                    TabIndex =3
                    Name ="04"
                    ControlSource ="04"
                    Format ="Standard"
                    EventProcPrefix ="Ctl04"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9105
                    Width =600
                    Height =225
                    TabIndex =4
                    Name ="05"
                    ControlSource ="05"
                    Format ="Standard"
                    EventProcPrefix ="Ctl05"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9765
                    Width =600
                    Height =225
                    TabIndex =5
                    Name ="06"
                    ControlSource ="06"
                    Format ="Standard"
                    EventProcPrefix ="Ctl06"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10425
                    Width =600
                    Height =225
                    TabIndex =6
                    Name ="07"
                    ControlSource ="07"
                    Format ="Standard"
                    EventProcPrefix ="Ctl07"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11085
                    Width =600
                    Height =225
                    TabIndex =7
                    Name ="08"
                    ControlSource ="08"
                    Format ="Standard"
                    EventProcPrefix ="Ctl08"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11745
                    Width =600
                    Height =225
                    TabIndex =8
                    Name ="09"
                    ControlSource ="09"
                    Format ="Standard"
                    EventProcPrefix ="Ctl09"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12405
                    Width =600
                    Height =225
                    TabIndex =9
                    Name ="10"
                    ControlSource ="10"
                    Format ="Standard"
                    EventProcPrefix ="Ctl10"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13065
                    Width =600
                    Height =225
                    TabIndex =10
                    Name ="11"
                    ControlSource ="11"
                    Format ="Standard"
                    EventProcPrefix ="Ctl11"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13725
                    Width =600
                    Height =225
                    TabIndex =11
                    Name ="12"
                    ControlSource ="12"
                    Format ="Standard"
                    EventProcPrefix ="Ctl12"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5565
                    Width =840
                    Height =225
                    TabIndex =12
                    Name ="Text77"
                    ControlSource ="TotalQty"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3960
                    Width =720
                    Height =225
                    TabIndex =13
                    Name ="Text95"
                    ControlSource ="SizeD"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4680
                    Width =840
                    Height =225
                    TabIndex =14
                    Name ="Text99"
                    ControlSource ="TotalDollars"
                    Format ="Standard"

                End
                Begin Line
                    Visible = NotDefault
                    Left =60
                    Top =240
                    Width =14280
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Width =2160
                    Height =225
                    ColumnWidth =2070
                    TabIndex =15
                    Name ="ClubName"
                    ControlSource ="ClubName"

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6465
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Text80"
                    ControlSource ="=Sum([01])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6465
                    LayoutCachedTop =60
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7125
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text81"
                    ControlSource ="=Sum([02])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7125
                    LayoutCachedTop =60
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7785
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Text82"
                    ControlSource ="=Sum([03])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7785
                    LayoutCachedTop =60
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8445
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Text83"
                    ControlSource ="=Sum([04])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8445
                    LayoutCachedTop =60
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9105
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Text84"
                    ControlSource ="=Sum([05])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9105
                    LayoutCachedTop =60
                    LayoutCachedWidth =9705
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9765
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text85"
                    ControlSource ="=Sum([06])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9765
                    LayoutCachedTop =60
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10425
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="Text86"
                    ControlSource ="=Sum([07])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10425
                    LayoutCachedTop =60
                    LayoutCachedWidth =11025
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11085
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="Text87"
                    ControlSource ="=Sum([08])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =11085
                    LayoutCachedTop =60
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11745
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Text88"
                    ControlSource ="=Sum([09])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =11745
                    LayoutCachedTop =60
                    LayoutCachedWidth =12345
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12405
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="Text89"
                    ControlSource ="=Sum([10])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =12405
                    LayoutCachedTop =60
                    LayoutCachedWidth =13005
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13065
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="Text90"
                    ControlSource ="=Sum([11])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =13065
                    LayoutCachedTop =60
                    LayoutCachedWidth =13665
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13725
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="Text91"
                    ControlSource ="=Sum([12])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =13725
                    LayoutCachedTop =60
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="Text92"
                    ControlSource ="=Sum([TotalQty])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4500
                    Top =60
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="Text100"
                    ControlSource ="=Sum([TotalDollars])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =2400
                    Height =285
                    FontSize =10
                    TabIndex =14
                    Name ="Text139"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =15
                    Name ="Text140"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =345
                End
                Begin Line
                    Top =420
                    Width =14280
                    Name ="Line142"
                    LayoutCachedTop =420
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =420
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Top =60
                    Width =3660
                    Name ="Text120"
                    ControlSource ="=\"Summary for \" & [CatDesc] & \" (\" & Count(*) & \" \" & IIf(Count(*)=1,\"det"
                        "ail record\",\"detail records\") & \")\""

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6465
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text121"
                    ControlSource ="=Sum([01])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7125
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Text122"
                    ControlSource ="=Sum([02])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7785
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Text123"
                    ControlSource ="=Sum([03])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8445
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Text124"
                    ControlSource ="=Sum([04])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9105
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text125"
                    ControlSource ="=Sum([05])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9765
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="Text126"
                    ControlSource ="=Sum([06])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =10425
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="Text127"
                    ControlSource ="=Sum([07])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11085
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Text128"
                    ControlSource ="=Sum([08])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11685
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="Text129"
                    ControlSource ="=Sum([09])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =12405
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="Text130"
                    ControlSource ="=Sum([10])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =13065
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="Text131"
                    ControlSource ="=Sum([11])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =13725
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="Text132"
                    ControlSource ="=Sum([12])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5820
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="Text133"
                    ControlSource ="=Sum([TotalQty])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Top =60
                    Width =1020
                    Height =300
                    FontWeight =700
                    TabIndex =14
                    Name ="Text134"
                    ControlSource ="=Sum([TotalDollars])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            Left =5580
                            Top =60
                            Width =195
                            Height =300
                            Name ="Label136"
                            Caption ="*"
                        End
                    End
                End
                Begin Line
                    Left =4620
                    Width =900
                    Name ="Line137"
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
                    Width =8640
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
                    Width =4740
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
                    Width =8640
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
                    Width =14280
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
                    Width =14280
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
            ForceNewPage =1
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6465
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="01 Grand Total Sum"
                    ControlSource ="=Sum([01])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl01_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label35"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7125
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="02 Grand Total Sum"
                    ControlSource ="=Sum([02])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl02_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7785
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="03 Grand Total Sum"
                    ControlSource ="=Sum([03])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl03_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label41"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8445
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="04 Grand Total Sum"
                    ControlSource ="=Sum([04])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl04_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9105
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="05 Grand Total Sum"
                    ControlSource ="=Sum([05])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl05_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label47"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9765
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="06 Grand Total Sum"
                    ControlSource ="=Sum([06])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl06_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label50"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =10425
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="07 Grand Total Sum"
                    ControlSource ="=Sum([07])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl07_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label53"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11085
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="08 Grand Total Sum"
                    ControlSource ="=Sum([08])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl08_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label56"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11745
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="09 Grand Total Sum"
                    ControlSource ="=Sum([09])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl09_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label59"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =12405
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="10 Grand Total Sum"
                    ControlSource ="=Sum([10])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl10_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label62"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =13065
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="11 Grand Total Sum"
                    ControlSource ="=Sum([11])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl11_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label65"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =13725
                    Top =120
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="12 Grand Total Sum"
                    ControlSource ="=Sum([12])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl12_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label68"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Top =120
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="Text79"
                    ControlSource ="=Sum([TotalQty])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3840
                    Top =120
                    Width =1680
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="Text101"
                    ControlSource ="=Sum([TotalDollars])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
            End
        End
    End
End
CodeBehindForm
' See "71 Bottle History By Club Month.cls"
