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
    Width =10380
    DatasheetFontHeight =10
    ItemSuffix =25
    Top =1275
    RecSrcDt = Begin
        0xb0c86033e7b3e340
    End
    RecordSource ="qryTodaysLiquor"
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
            Height =1200
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2970
                    Top =765
                    Width =900
                    Height =315
                    Name ="Counter"
                    ControlSource ="Counter"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =2970
                            Top =420
                            Width =900
                            Height =360
                            FontSize =10
                            Name ="Counter_Label"
                            Caption ="Counter"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3870
                    Top =765
                    Width =1065
                    Height =315
                    TabIndex =1
                    Name ="ClubsGross"
                    ControlSource ="Clubs"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =3870
                            Top =420
                            Width =1065
                            Height =360
                            FontSize =10
                            Name ="ClubsGross_Label"
                            Caption ="Clubs Gross"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4950
                    Top =765
                    Width =900
                    Height =315
                    TabIndex =2
                    Name ="ClubsDiscount"
                    ControlSource ="Purchases"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =4950
                            Top =420
                            Width =915
                            Height =360
                            FontSize =10
                            Name ="ClubsDiscount_Label"
                            Caption ="Purchases"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5850
                    Top =765
                    Width =1020
                    Height =315
                    TabIndex =3
                    Name ="ClubsNet"
                    ControlSource ="TransfersIN"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =5850
                            Top =420
                            Width =1020
                            Height =360
                            FontSize =10
                            Name ="ClubsNet_Label"
                            Caption ="In Transfer"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8070
                    Top =765
                    Width =1200
                    Height =315
                    TabIndex =4
                    Name ="OLCCDeposit"
                    ControlSource ="Adjustments"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =8070
                            Top =420
                            Width =1200
                            Height =360
                            FontSize =10
                            Name ="OLCCDeposit_Label"
                            Caption ="Adjustments"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9270
                    Top =765
                    Width =1020
                    Height =315
                    TabIndex =5
                    Name ="Cards"
                    ControlSource ="CurrentLiquorValue"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =9270
                            Top =420
                            Width =1020
                            Height =360
                            FontSize =10
                            Name ="Cards_Label"
                            Caption ="Current"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =270
                    Top =765
                    Width =1545
                    Height =315
                    FontWeight =700
                    TabIndex =6
                    Name ="CashRequired"
                    ControlSource ="Today"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =270
                            Top =420
                            Width =1545
                            Height =360
                            FontSize =10
                            Name ="CashRequired_Label"
                            Caption ="Today's Liquor "
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1830
                    Top =765
                    Width =1155
                    Height =315
                    TabIndex =7
                    Name ="Text20"
                    ControlSource ="LastNight"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =1830
                            Top =420
                            Width =1155
                            Height =360
                            FontSize =10
                            Name ="Label21"
                            Caption ="Last Night"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6870
                    Top =765
                    Width =1200
                    Height =315
                    TabIndex =8
                    Name ="Text22"
                    ControlSource ="TransfersOUT"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =2
                            Left =6870
                            Top =420
                            Width =1215
                            Height =360
                            FontSize =10
                            Name ="Label23"
                            Caption ="Out Transfer"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =180
                    Top =60
                    Width =10200
                    Height =1080
                    FontSize =10
                    Name ="Label24"
                    Caption ="Today's Liquor Recap"
                    FontName ="Arial Narrow"
                End
            End
        End
    End
End
