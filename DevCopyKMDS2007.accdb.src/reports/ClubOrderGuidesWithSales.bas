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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =201
    Left =1140
    Top =1980
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0xdb3fc60552ade540
    End
    RecordSource ="qryClubOrderGuidesWithActualSales"
    Caption ="Club Order Guide"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
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
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2160
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9840
                    Top =780
                    Width =720
                    ColumnOrder =0
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Left =7080
                    Top =420
                    Width =3420
                    Height =270
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text77"
                    ControlSource ="ClubName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =6960
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label69"
                            Caption ="    Ship to / Sold to"
                        End
                    End
                End
                Begin TextBox
                    Left =7080
                    Top =660
                    Width =3420
                    Height =270
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =7080
                    Top =900
                    Width =3420
                    Height =270
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =7080
                    Top =1200
                    Width =3420
                    Height =270
                    ColumnOrder =4
                    TabIndex =4
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =7080
                    Top =1440
                    Width =3420
                    Height =270
                    ColumnOrder =5
                    TabIndex =5
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin Rectangle
                    Left =6960
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =7620
                    Top =1740
                    Width =1260
                    Height =270
                    ColumnOrder =6
                    TabIndex =6
                    Name ="Text148"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =7080
                    Top =1740
                    Width =480
                    Height =300
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text147"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =8880
                    Top =1740
                    Width =420
                    Height =270
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =9300
                    Top =1740
                    Width =1260
                    Height =300
                    ColumnOrder =9
                    TabIndex =9
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =240
                    Top =420
                    Width =3420
                    Height =270
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =10
                    Name ="Text90"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =120
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label91"
                            Caption ="    Shipped from "
                        End
                    End
                End
                Begin TextBox
                    Left =240
                    Top =660
                    Width =3420
                    Height =270
                    ColumnOrder =11
                    TabIndex =11
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =240
                    Top =900
                    Width =3420
                    Height =270
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =240
                    Top =1200
                    Width =3420
                    Height =270
                    ColumnOrder =13
                    TabIndex =13
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =240
                    Top =1440
                    Width =3420
                    Height =270
                    ColumnOrder =14
                    TabIndex =14
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin Rectangle
                    Left =120
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box97"
                End
                Begin TextBox
                    Left =780
                    Top =1740
                    Width =1260
                    Height =300
                    ColumnOrder =15
                    TabIndex =15
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =1980
                    Top =1740
                    Width =420
                    Height =270
                    ColumnOrder =16
                    TabIndex =16
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2400
                    Top =1740
                    Width =1260
                    Height =300
                    ColumnOrder =17
                    TabIndex =17
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =240
                    Top =1740
                    Width =480
                    Height =300
                    ColumnOrder =18
                    TabIndex =18
                    Name ="Text146"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3900
                    Top =60
                    Width =3000
                    Height =540
                    ColumnOrder =19
                    FontSize =20
                    FontWeight =700
                    TabIndex =19
                    ForeColor =8388608
                    Name ="Text175"
                    ControlSource ="=Format(Now(),\"mmm-yy\")"
                    FontName ="Trebuchet MS"

                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =660
                    Width =3000
                    Height =300
                    FontSize =10
                    Name ="Label176"
                    Caption ="Prices with * are on sale this month"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =1080
                    Width =3000
                    Height =300
                    FontSize =10
                    Name ="Label177"
                    Caption ="Net Price/oz reflects licensee disc't."
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =1500
                    Width =3000
                    Height =600
                    FontSize =10
                    Name ="Label178"
                    Caption ="Prices are set by the OLCC and may change next month."
                End
            End
        End
        Begin PageHeader
            Height =780
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =420
                    Width =1020
                    Height =300
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9660
                            Top =120
                            Width =1020
                            Height =300
                            BackColor =16053492
                            Name ="lblPageNumber"
                            Caption ="Page "
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Top =120
                    Width =6960
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =6960
                    Top =120
                    Width =2640
                    Height =600
                    FontSize =20
                    Name ="Label50"
                    Caption ="Order Guide"
                    FontName ="Trebuchet MS"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =540
            Name ="GroupHeader0"
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =600
                    Width =720
                    Height =540
                    FontSize =10
                    Name ="Label83"
                    Caption ="Item"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1320
                    Width =2520
                    Height =540
                    FontSize =10
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6840
                    Width =720
                    Height =540
                    FontSize =10
                    Name ="Label85"
                    Caption ="Week 1"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4740
                    Width =1020
                    Height =540
                    FontSize =10
                    Name ="Label145"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6240
                    Width =600
                    Height =540
                    FontSize =10
                    Name ="Label157"
                    Caption ="Par"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3840
                    Width =900
                    Height =540
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text180"
                    ControlSource ="=\"Net $/ \" & [OZColumn] & \" oz\""
                    FontName ="Times New Roman"

                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7560
                    Width =780
                    Height =540
                    FontSize =10
                    Name ="Label163"
                    Caption ="Week 2"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8340
                    Width =780
                    Height =540
                    FontSize =10
                    Name ="Label164"
                    Caption ="Week 3"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9120
                    Width =780
                    Height =540
                    FontSize =10
                    Name ="Label165"
                    Caption ="Week 4"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9900
                    Width =780
                    Height =540
                    FontSize =10
                    Name ="Label166"
                    Caption ="Week 5"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =5400
                    Height =360
                    TabIndex =6
                    Name ="Text141"

                End
                Begin TextBox
                    TextAlign =2
                    Left =6300
                    Top =60
                    Width =420
                    FontSize =6
                    TabIndex =5
                    Name ="Par"
                    ControlSource ="qryClubOrderGuidesReport.Par"
                    Format ="#.00;#.00;\"\";\"\""

                    LayoutCachedLeft =6300
                    LayoutCachedTop =60
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =2
                    Left =660
                    Top =60
                    Width =600
                    Height =225
                    FontWeight =700
                    Name ="ItemID"
                    ControlSource ="STRIPPED"

                End
                Begin TextBox
                    TextAlign =1
                    Left =1440
                    Top =60
                    Width =2460
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="LineDesc"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =6240
                    Width =0
                    Height =420
                    Name ="Line91"
                End
                Begin Line
                    Left =600
                    Width =0
                    Height =420
                    Name ="Line93"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    BackStyle =1
                    Left =3900
                    Top =60
                    Width =780
                    TabIndex =2
                    Name ="PricePer125"
                    ControlSource ="qryClubOrderGuidesReport.PricePerShot"
                    Format ="Fixed"

                End
                Begin Line
                    Left =1320
                    Width =0
                    Height =420
                    Name ="Line134"
                End
                Begin Line
                    Left =4740
                    Width =0
                    Height =420
                    Name ="Line135"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4800
                    Top =60
                    Width =660
                    TabIndex =3
                    Name ="TodaysPrice"
                    ControlSource ="TodaysPrice"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =120
                    TabIndex =4
                    Name ="Sprice"
                    ControlSource ="Sprice"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =600
                    Top =420
                    Width =10080
                    Name ="Line146"
                End
                Begin Line
                    Left =3840
                    Width =0
                    Height =420
                    Name ="Line162"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5760
                    Width =0
                    Height =420
                    Name ="Line167"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =6840
                    Width =0
                    Height =420
                    Name ="Line168"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7560
                    Width =0
                    Height =420
                    Name ="Line169"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =8340
                    Width =0
                    Height =420
                    Name ="Line170"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9120
                    Width =0
                    Height =420
                    Name ="Line171"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9900
                    Width =0
                    Height =420
                    Name ="Line172"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10680
                    Width =0
                    Height =420
                    Name ="Line173"
                End
                Begin TextBox
                    TextAlign =2
                    Left =7080
                    Top =60
                    Width =300
                    TabIndex =7
                    Name ="Text179"
                    ControlSource ="W1"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =7800
                    Top =60
                    Width =300
                    TabIndex =8
                    Name ="Text181"
                    ControlSource ="W2"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =8580
                    Top =60
                    Width =300
                    TabIndex =9
                    Name ="Text182"
                    ControlSource ="W3"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =9360
                    Top =60
                    Width =300
                    TabIndex =10
                    Name ="Text183"
                    ControlSource ="W4"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =10140
                    Top =60
                    Width =300
                    TabIndex =11
                    Name ="Text184"
                    ControlSource ="W5"
                    Format ="#;#;\"\";\"\""

                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =600
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =600
                    Width =9480
                    Name ="Line109"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7560
                    Width =0
                    Height =420
                    Name ="Line185"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =8340
                    Width =0
                    Height =420
                    Name ="Line186"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9120
                    Width =0
                    Height =420
                    Name ="Line187"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9900
                    Width =0
                    Height =420
                    Name ="Line188"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10680
                    Width =0
                    Height =420
                    Name ="Line189"
                End
                Begin TextBox
                    TextAlign =2
                    Left =7080
                    Top =60
                    Width =300
                    Name ="Text190"
                    ControlSource ="=Sum([W1])"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =7800
                    Top =60
                    Width =300
                    TabIndex =1
                    Name ="Text191"
                    ControlSource ="=Sum([W2])"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =8580
                    Top =60
                    Width =300
                    TabIndex =2
                    Name ="Text192"
                    ControlSource ="=Sum([W3])"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =9360
                    Top =60
                    Width =300
                    TabIndex =3
                    Name ="Text193"
                    ControlSource ="=Sum([W4])"
                    Format ="#;#;\"\";\"\""

                End
                Begin TextBox
                    TextAlign =2
                    Left =10140
                    Top =60
                    Width =300
                    TabIndex =4
                    Name ="Text194"
                    ControlSource ="=Sum([W5])"
                    Format ="#;#;\"\";\"\""

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =6840
                    Width =0
                    Height =420
                    Name ="Line196"
                End
                Begin Label
                    Left =600
                    Width =5700
                    Height =240
                    FontSize =8
                    FontWeight =400
                    Name ="Label197"
                    Caption =" Note: Items may have been purchased that are not on this order guide...."
                End
                Begin Line
                    Left =6840
                    Top =420
                    Width =3840
                    Name ="Line198"
                End
                Begin Label
                    Left =1140
                    Top =180
                    Width =3975
                    Height =240
                    FontSize =8
                    FontWeight =400
                    Name ="Label199"
                    Caption =" We can supply you with a full purchase history by month... "
                End
                Begin Label
                    Left =1140
                    Top =360
                    Width =4020
                    Height =240
                    FontSize =8
                    FontWeight =400
                    Name ="Label200"
                    Caption =" Request report #24 Club History by Bottle by Month"
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrderGuidesWithSales.cls"
