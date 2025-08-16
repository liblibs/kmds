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
    Width =10755
    DatasheetFontHeight =10
    ItemSuffix =161
    Left =1020
    Top =120
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xc47299ef51ade540
    End
    RecordSource ="qryClubOrderGuidesReportAndSysfile"
    Caption ="Club Order Guide"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    ShowPageMargins =0
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
            Height =2040
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9840
                    Top =360
                    Width =720
                    ColumnOrder =0
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Left =7140
                    Top =360
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
                            Left =7020
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label69"
                            Caption ="    Ship to / Sold to"
                        End
                    End
                End
                Begin TextBox
                    Left =7140
                    Top =600
                    Width =3420
                    Height =270
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =7140
                    Top =840
                    Width =3420
                    Height =270
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =7140
                    Top =1140
                    Width =3420
                    Height =270
                    ColumnOrder =4
                    TabIndex =4
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =7140
                    Top =1380
                    Width =3420
                    Height =270
                    ColumnOrder =5
                    TabIndex =5
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin Rectangle
                    Left =7020
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =7680
                    Top =1680
                    Width =1260
                    Height =270
                    ColumnOrder =6
                    TabIndex =6
                    Name ="Text148"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =8880
                    Top =1680
                    Width =420
                    Height =270
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =9300
                    Top =1680
                    Width =1260
                    Height =300
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =7140
                    Top =1680
                    Width =480
                    Height =300
                    ColumnOrder =9
                    TabIndex =9
                    Name ="Text147"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =180
                    Top =360
                    Width =3480
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
                            Left =60
                            Width =3780
                            Height =300
                            BackColor =16053492
                            Name ="Label91"
                            Caption ="    Shipped from "
                        End
                    End
                End
                Begin TextBox
                    Left =180
                    Top =600
                    Width =3480
                    Height =270
                    ColumnOrder =11
                    TabIndex =11
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =180
                    Top =840
                    Width =3480
                    Height =270
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =180
                    Top =1140
                    Width =3480
                    Height =270
                    ColumnOrder =13
                    TabIndex =13
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =180
                    Top =1380
                    Width =3480
                    Height =270
                    ColumnOrder =14
                    TabIndex =14
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin Rectangle
                    Left =60
                    Width =3780
                    Height =2040
                    Name ="Box97"
                End
                Begin TextBox
                    Left =780
                    Top =1620
                    Width =1260
                    Height =300
                    ColumnOrder =15
                    TabIndex =15
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =1920
                    Top =1620
                    Width =480
                    Height =270
                    ColumnOrder =16
                    TabIndex =16
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2340
                    Top =1620
                    Width =1320
                    Height =300
                    ColumnOrder =17
                    TabIndex =17
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =240
                    Top =1620
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
                    Width =3060
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
                    Top =1140
                    Width =3060
                    Height =300
                    FontSize =10
                    Name ="Label176"
                    Caption ="Prices with * are on sale this month"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =1440
                    Width =3060
                    Height =600
                    FontSize =10
                    Name ="Label178"
                    Caption ="Prices are set by the OLCC and may change next month."
                End
            End
        End
        Begin PageHeader
            Height =685
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9720
                    Top =345
                    Width =1020
                    Height =300
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =345
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9720
                            Top =45
                            Width =1020
                            Height =300
                            BackColor =16053492
                            Name ="lblPageNumber"
                            Caption ="Page "
                            LayoutCachedLeft =9720
                            LayoutCachedTop =45
                            LayoutCachedWidth =10740
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Top =45
                    Width =6960
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                    LayoutCachedTop =45
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =645
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =6960
                    Top =45
                    Width =2640
                    Height =600
                    FontSize =20
                    Name ="Label50"
                    Caption ="Order Guide"
                    FontName ="Trebuchet MS"
                    LayoutCachedLeft =6960
                    LayoutCachedTop =45
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =645
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =360
            Name ="GroupHeader0"
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1020
                    Top =60
                    Width =720
                    Height =300
                    FontSize =10
                    Name ="Label83"
                    Caption ="Item"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1740
                    Top =60
                    Width =2160
                    Height =300
                    FontSize =10
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label85"
                    Caption ="Qty"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label177"
                    Caption ="Size"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4500
                    Top =60
                    Width =840
                    Height =300
                    FontSize =10
                    Name ="Label145"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =660
                    Top =60
                    Width =360
                    Height =300
                    FontSize =10
                    Name ="Label157"
                    Caption ="Par"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =720
                    Top =60
                    Width =240
                    FontSize =6
                    TabIndex =5
                    Name ="Par"
                    ControlSource ="qryClubOrderGuidesReport.Par"
                    Format ="#.0;#.0;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1800
                    Top =60
                    Width =3480
                    Height =300
                    TabIndex =6
                    BackColor =16053492
                    Name ="Text141"

                End
                Begin TextBox
                    TextFontCharSet =161
                    TextAlign =3
                    Left =795
                    Top =60
                    Width =885
                    Height =225
                    FontWeight =700
                    Name ="ItemID"
                    ControlSource ="STRIPPED"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =795
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =1
                    Left =1800
                    Top =60
                    Width =2100
                    Height =225
                    TabIndex =1
                    BackColor =16053492
                    Name ="Description"
                    ControlSource ="LineDesc"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =5340
                    Width =0
                    Height =360
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =0
                    Height =360
                    Name ="Line92"
                End
                Begin Line
                    Left =660
                    Width =0
                    Height =360
                    Name ="Line93"
                End
                Begin TextBox
                    BackStyle =1
                    Left =3960
                    Top =60
                    Width =480
                    Height =225
                    TabIndex =2
                    BackColor =16053492
                    Name ="SizeID"
                    ControlSource ="SizeDesc"

                End
                Begin Line
                    Left =1740
                    Width =0
                    Height =360
                    Name ="Line134"
                End
                Begin Line
                    Left =4500
                    Width =0
                    Height =360
                    Name ="Line135"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =735
                    TabIndex =3
                    BackColor =16053492
                    Name ="TodaysPrice"
                    ControlSource ="TodaysPrice"
                    Format ="Standard"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =60
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5220
                    Top =60
                    Width =120
                    TabIndex =4
                    BackColor =16053492
                    Name ="Sprice"
                    ControlSource ="Sprice"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Top =360
                    Width =1680
                    Name ="Line146"
                End
                Begin Line
                    Left =1020
                    Width =0
                    Height =360
                    Name ="Line158"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "ClubOrderGuides.cls"
