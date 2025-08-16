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
    Width =19140
    DatasheetFontHeight =10
    ItemSuffix =178
    Left =2145
    Top =690
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x44cd560152ade540
    End
    RecordSource ="qryClubOrderGuidesByCategoryReportAndSysfile"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Category"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =600
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7650
                    Top =60
                    Width =3105
                    Height =540
                    FontSize =20
                    Name ="Label50"
                    Caption ="Order Guide"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =7560
                    Height =510
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7620
                    Top =240
                    Width =720
                    ColumnOrder =1
                    TabIndex =1
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
            End
        End
        Begin PageHeader
            Height =2100
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9420
                    Top =360
                    Width =1380
                    Height =300
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9420
                            Top =60
                            Width =1380
                            Height =300
                            BackColor =16053492
                            Name ="lblPageNumber"
                            Caption ="Page Number"
                        End
                    End
                End
                Begin TextBox
                    Left =5580
                    Top =420
                    Width =3420
                    Height =270
                    TabIndex =1
                    Name ="Text77"
                    ControlSource ="ClubName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =5460
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
                    Left =5580
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =5580
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =5580
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =5580
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =5
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin Rectangle
                    Left =5460
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =180
                    Top =420
                    Width =3420
                    Height =270
                    FontWeight =700
                    TabIndex =6
                    Name ="Text90"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =60
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
                    Left =180
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =180
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =180
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =180
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin Rectangle
                    Left =60
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box97"
                End
                Begin TextBox
                    Left =6120
                    Top =1740
                    Width =1260
                    Height =270
                    TabIndex =11
                    Name ="Text148"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =7320
                    Top =1740
                    Width =420
                    Height =270
                    TabIndex =12
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =7740
                    Top =1740
                    Width =1260
                    Height =300
                    TabIndex =13
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =5580
                    Top =1740
                    Width =480
                    Height =300
                    TabIndex =14
                    Name ="Text147"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =720
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =15
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =1920
                    Top =1680
                    Width =420
                    Height =270
                    TabIndex =16
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2340
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =17
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =180
                    Top =1680
                    Width =480
                    Height =300
                    TabIndex =18
                    Name ="Text146"
                    ControlSource ="=\"Voice: \""

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =420
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
                    Name ="Label158"
                    Caption ="Item"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1740
                    Top =60
                    Width =2160
                    Height =300
                    FontSize =10
                    Name ="Label159"
                    Caption ="Description"
                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label160"
                    Caption ="Qty"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =660
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =60
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label161"
                    Caption ="Size"
                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4500
                    Top =60
                    Width =840
                    Height =300
                    FontSize =10
                    Name ="Label162"
                    Caption ="Price"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
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
                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =360
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =4560
                    FontWeight =700
                    Name ="Category"
                    ControlSource ="Category"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =0
                    Height =420
                    Name ="Line150"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5340
                    Width =0
                    Height =420
                    Name ="Line151"
                    LayoutCachedLeft =5340
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Top =360
                    Width =5280
                    Name ="Line152"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =5280
                    Name ="Line153"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =5340
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4500
                    Top =360
                    Width =0
                    Height =60
                    Name ="Line154"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =360
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =660
                    Top =360
                    Width =0
                    Height =60
                    Name ="Line156"
                    LayoutCachedLeft =660
                    LayoutCachedTop =360
                    LayoutCachedWidth =660
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1740
                    Top =360
                    Width =0
                    Height =60
                    Name ="Line157"
                    LayoutCachedLeft =1740
                    LayoutCachedTop =360
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1020
                    Top =360
                    Width =0
                    Height =60
                    Name ="Line164"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =360
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =420
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
                    Name ="Par"
                    ControlSource ="qryClubOrderGuidesReportByCategory.Par"
                    Format ="#;#;\"\";\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =960
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1800
                    Top =60
                    Width =3480
                    Height =300
                    TabIndex =1
                    BackColor =16053492
                    Name ="Text141"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    Left =840
                    Top =60
                    Width =840
                    Height =225
                    FontWeight =700
                    TabIndex =2
                    Name ="ItemID"
                    ControlSource ="STRIPPED"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =840
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
                    TabIndex =3
                    BackColor =16053492
                    Name ="Description"
                    ControlSource ="LineDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =285
                End
                Begin Line
                    Left =5340
                    Width =0
                    Height =360
                    Name ="Line91"
                    LayoutCachedLeft =5340
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =0
                    Height =360
                    Name ="Line92"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =60
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =660
                    Width =0
                    Height =360
                    Name ="Line93"
                    LayoutCachedLeft =660
                    LayoutCachedWidth =660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    BackStyle =1
                    Left =3960
                    Top =60
                    Width =480
                    Height =225
                    TabIndex =4
                    BackColor =16053492
                    Name ="SizeID"
                    ControlSource ="SizeDesc"

                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =285
                End
                Begin Line
                    Left =1740
                    Width =0
                    Height =360
                    Name ="Line134"
                    LayoutCachedLeft =1740
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =4500
                    Width =0
                    Height =360
                    Name ="Line135"
                    LayoutCachedLeft =4500
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =735
                    TabIndex =5
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
                    TabIndex =6
                    BackColor =16053492
                    Name ="Sprice"
                    ControlSource ="Sprice"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =300
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Top =360
                    Width =5280
                    Name ="Line146"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =1020
                    Width =0
                    Height =360
                    Name ="Line158"
                    LayoutCachedLeft =1020
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =360
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
            Visible = NotDefault
            CanGrow = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "ClubOrderGuidesByCategory.cls"
