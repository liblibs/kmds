Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =49
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =26
    Top =795
    RecSrcDt = Begin
        0x169e6f3449bce540
    End
    RecordSource ="qryOLCCDepositRecap"
    Caption ="qryOLCCDepositRecap"
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
        Begin Section
            KeepTogether = NotDefault
            Height =780
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =360
                    Width =990
                    Height =300
                    Name ="Counter"
                    ControlSource ="Counter"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =1020
                            Width =990
                            Height =660
                            FontSize =8
                            Name ="Counter_Label"
                            Caption ="#1 Counter"
                            FontName ="Arial Narrow"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =360
                    Width =1020
                    Height =300
                    TabIndex =1
                    Name ="ClubsGross"
                    ControlSource ="ClubsGross"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =360
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =2040
                            Width =1020
                            Height =660
                            FontSize =8
                            Name ="ClubsGross_Label"
                            Caption ="#3 Licensee"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =2040
                            LayoutCachedWidth =3060
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3060
                    Top =360
                    Width =960
                    Height =300
                    TabIndex =2
                    Name ="ClubsDiscount"
                    ControlSource ="ClubsDiscount"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3060
                    LayoutCachedTop =360
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =3075
                            Width =960
                            Height =660
                            FontSize =8
                            Name ="ClubsDiscount_Label"
                            Caption ="Discount"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =3075
                            LayoutCachedWidth =4035
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4020
                    Top =360
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="ClubsNet"
                    ControlSource ="ClubsNet"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =360
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =4020
                            Width =1020
                            Height =660
                            FontSize =8
                            Name ="ClubsNet_Label"
                            Caption ="Net Clubs"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =4020
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5040
                    Top =360
                    Width =870
                    Height =300
                    TabIndex =4
                    Name ="OLCCDeposit"
                    ControlSource ="OLCCDeposit"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =360
                    LayoutCachedWidth =5910
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =5040
                            Width =870
                            Height =660
                            FontSize =8
                            Name ="OLCCDeposit_Label"
                            Caption =" Net Sales"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =5040
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7800
                    Top =360
                    Width =1020
                    Height =300
                    TabIndex =5
                    Name ="Checks"
                    ControlSource ="Checks"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =360
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =7800
                            Width =1020
                            Height =660
                            FontSize =8
                            Name ="Checks_Label"
                            Caption ="Checks"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =7800
                            LayoutCachedWidth =8820
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =360
                    Width =960
                    Height =300
                    TabIndex =6
                    Name ="Cards"
                    ControlSource ="Cards"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =360
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =6840
                            Width =960
                            Height =660
                            FontSize =8
                            Name ="Cards_Label"
                            Caption ="#15 Cards"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =6840
                            LayoutCachedWidth =7800
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8835
                    Top =360
                    Width =945
                    Height =300
                    FontWeight =700
                    TabIndex =7
                    Name ="CashRequired"
                    ControlSource ="CashRequired"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8835
                    LayoutCachedTop =360
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =8820
                            Width =960
                            Height =660
                            FontSize =8
                            Name ="CashRequired_Label"
                            Caption ="Cash "
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =8820
                            LayoutCachedWidth =9780
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Top =360
                    Width =1020
                    Height =300
                    TabIndex =8
                    Name ="Text20"
                    ControlSource ="=[Counter]+[ClubsGross]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedTop =360
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Width =1020
                            Height =660
                            FontSize =8
                            Name ="Label21"
                            Caption ="Retail"
                            FontName ="Arial Narrow"
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9780
                    Top =360
                    Width =930
                    Height =300
                    FontWeight =700
                    TabIndex =9
                    Name ="Text22"
                    ControlSource ="Deposit"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =360
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =9780
                            Width =930
                            Height =660
                            FontSize =8
                            Name ="Label23"
                            Caption ="#16 Deposit"
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =9780
                            LayoutCachedWidth =10710
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5880
                    Top =360
                    Width =960
                    Height =300
                    TabIndex =10
                    Name ="Text24"
                    ControlSource ="EFT"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =360
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =5880
                            Width =960
                            Height =660
                            FontSize =8
                            Name ="Label25"
                            Caption ="#14 EFT "
                            FontName ="Arial Narrow"
                            LayoutCachedLeft =5880
                            LayoutCachedWidth =6840
                            LayoutCachedHeight =660
                        End
                    End
                End
            End
        End
    End
End
