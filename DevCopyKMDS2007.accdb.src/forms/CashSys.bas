Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9915
    DatasheetFontHeight =10
    ItemSuffix =62
    Top =600
    Bottom =600
    RecSrcDt = Begin
        0xadb6b003fa92e240
    End
    RecordSource ="CashSys"
    Caption ="CashSys"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =5100
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =135
                    Width =9915
                    Height =4965
                    Name ="TabCtl56"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =540
                            Width =9645
                            Height =4425
                            Name ="Page57"
                            Begin
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =780
                                    Name ="ID"
                                    ControlSource ="ID"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1875
                                            Top =780
                                            Width =240
                                            Height =240
                                            Name ="Label1"
                                            Caption ="ID"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =1140
                                    Name ="ShiftCounter"
                                    ControlSource ="ShiftCounter"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1125
                                            Top =1140
                                            Width =990
                                            Height =240
                                            Name ="Label3"
                                            Caption ="ShiftCounter"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =1500
                                    Width =2490
                                    Name ="ComputerName"
                                    ControlSource ="ComputerName"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =930
                                            Top =1500
                                            Width =1185
                                            Height =240
                                            Name ="Label5"
                                            Caption ="ComputerName"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =1860
                                    Width =2490
                                    Name ="Description"
                                    ControlSource ="Description"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1245
                                            Top =1860
                                            Width =870
                                            Height =240
                                            Name ="Label7"
                                            Caption ="Description"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =2220
                                    Width =2490
                                    Name ="ReportPort"
                                    ControlSource ="ReportPort"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1245
                                            Top =2220
                                            Width =870
                                            Height =240
                                            Name ="Label9"
                                            Caption ="ReportPort"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =2580
                                    Width =2490
                                    Name ="DisplayPort"
                                    ControlSource ="DisplayPort"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1230
                                            Top =2580
                                            Width =885
                                            Height =240
                                            Name ="Label11"
                                            Caption ="DisplayPort"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =2940
                                    Width =2490
                                    Name ="DisplayDefault"
                                    ControlSource ="DisplayDefault"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1005
                                            Top =2940
                                            Width =1110
                                            Height =240
                                            Name ="Label13"
                                            Caption ="DisplayDefault"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =3300
                                    Width =2490
                                    Name ="PrinterPort"
                                    ControlSource ="PrinterPort"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1260
                                            Top =3300
                                            Width =855
                                            Height =240
                                            Name ="Label15"
                                            Caption ="PrinterPort"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =3660
                                    Width =2490
                                    Name ="ReportsDir"
                                    ControlSource ="ReportsDir"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1275
                                            Top =3660
                                            Width =840
                                            Height =240
                                            Name ="Label17"
                                            Caption ="ReportsDir"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =2220
                                    Top =4020
                                    Width =2490
                                    Name ="InvoiceTrailer"
                                    ControlSource ="InvoiceTrailer"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =1065
                                            Top =4020
                                            Width =1050
                                            Height =240
                                            Name ="Label19"
                                            Caption ="InvoiceTrailer"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =2220
                                    Top =4380
                                    Name ="ScannerComPort"
                                    ControlSource ="ScannerComPort"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =3840
                                            Top =4380
                                            Width =5010
                                            Height =240
                                            Name ="Label61"
                                            Caption ="Enter \"0\" if there is no scanner OR scanner has a keyboard interface"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =2220
                                    Top =4680
                                    Name ="ScannerSettings"
                                    ControlSource ="ScannerSettings"
                                    DefaultValue ="\"9600,N,8,1\""
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =810
                                            Top =4680
                                            Width =1305
                                            Height =240
                                            Name ="Label60"
                                            Caption ="ScannerSettings:"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =780
                                    Top =4380
                                    Width =1335
                                    Height =240
                                    Name ="Label59"
                                    Caption ="ScannerComPort:"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =540
                            Width =9645
                            Height =4425
                            Name ="Page58"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =710
                                    Name ="AllowExactCh"
                                    ControlSource ="AllowExactCh"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =710
                                            Width =1050
                                            Height =240
                                            Name ="Label21"
                                            Caption ="AllowExactCh"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =1070
                                    Name ="AllowEscape"
                                    ControlSource ="AllowEscape"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =1070
                                            Width =960
                                            Height =240
                                            Name ="Label23"
                                            Caption ="AllowEscape"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =1430
                                    Name ="Reprintnegsales"
                                    ControlSource ="Reprintnegsales"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =1430
                                            Width =1230
                                            Height =240
                                            Name ="Label25"
                                            Caption ="Reprintnegsales"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =1790
                                    Name ="EpsonSlitIns"
                                    ControlSource ="EpsonSlitIns"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =1790
                                            Width =945
                                            Height =240
                                            Name ="Label27"
                                            Caption ="EpsonSlitIns"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =2150
                                    Name ="PrintReceipt"
                                    ControlSource ="PrintReceipt"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =2150
                                            Width =945
                                            Height =240
                                            Name ="Label29"
                                            Caption ="PrintReceipt"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =2510
                                    Name ="EpsonPrinterInstalled"
                                    ControlSource ="EpsonPrinterInstalled"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =2510
                                            Width =1605
                                            Height =240
                                            Name ="Label31"
                                            Caption ="EpsonPrinterInstalled"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2300
                                    Top =2870
                                    Name ="EpsonPrintertSlit"
                                    ControlSource ="EpsonPrintertSlit"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =500
                                            Top =2870
                                            Width =1260
                                            Height =240
                                            Name ="Label33"
                                            Caption ="EpsonPrintertSlit"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2280
                                    Top =3300
                                    Name ="DropAlertAmt"
                                    ControlSource ="DropAlertAmt"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =480
                                            Top =3300
                                            Width =1050
                                            Height =240
                                            Name ="Label35"
                                            Caption ="DropAlertAmt"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2280
                                    Top =3660
                                    Name ="LastInvoiceNo"
                                    ControlSource ="LastInvoiceNo"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =480
                                            Top =3660
                                            Width =1095
                                            Height =240
                                            Name ="Label37"
                                            Caption ="LastInvoiceNo"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2280
                                    Top =4020
                                    Name ="ReceiptSpaces"
                                    ControlSource ="ReceiptSpaces"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =480
                                            Top =4020
                                            Width =1125
                                            Height =240
                                            Name ="Label39"
                                            Caption ="ReceiptSpaces"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1260
                    Top =60
                    Width =2040
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =60
                            Width =840
                            Height =240
                            Name ="Label53"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6360
                    Top =60
                    Width =2040
                    TabIndex =1
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5280
                            Top =60
                            Width =975
                            Height =240
                            Name ="Label55"
                            Caption ="ChangeDate"
                        End
                    End
                End
            End
        End
    End
End
