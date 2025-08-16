Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15545
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =2010
    Top =1125
    Right =12600
    Bottom =11355
    TimerInterval =1000
    RecSrcDt = Begin
        0x29d86c428be7e440
    End
    RecordSource ="qryClubOrderCSV"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =161
            FontSize =10
            BorderColor =5137573
            ForeColor =5137573
            FontName ="Franklin Gothic Book"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =5137573
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =5137573
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =5137573
        End
        Begin CommandButton
            TextFontCharSet =161
            FontWeight =400
            ForeColor =3160910
            FontName ="Franklin Gothic Book"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =5137573
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =5137573
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =5137573
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =5137573
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =5137573
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =5137573
        End
        Begin ToggleButton
            TextFontCharSet =161
            FontWeight =400
            ForeColor =3160910
            FontName ="Franklin Gothic Book"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =5137573
            LabelX =-1800
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =1680
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2760
                    Top =570
                    Width =1320
                    Height =315
                    ColumnOrder =4
                    Name ="txtInterval"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =570
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =885
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =2
                            Left =2760
                            Top =180
                            Width =1320
                            Height =300
                            Name ="Label20"
                            Caption ="Next poll ..."
                            LayoutCachedLeft =2760
                            LayoutCachedTop =180
                            LayoutCachedWidth =4080
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =570
                    Width =2400
                    Height =315
                    ColumnOrder =5
                    TabIndex =1
                    Name ="txtTime"

                    LayoutCachedLeft =180
                    LayoutCachedTop =570
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =885
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =180
                            Top =180
                            Width =2400
                            Height =300
                            Name ="Label18"
                            Caption ="Current time:"
                            LayoutCachedLeft =180
                            LayoutCachedTop =180
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8340
                    Top =570
                    Width =2040
                    Height =315
                    ColumnOrder =6
                    TabIndex =2
                    Name ="txtIntervalSet"
                    Format ="Fixed"
                    ValidationRule ="Between 1 And 120"
                    ValidationText ="Must be between 1 and 120 minutes.."
                    DefaultValue ="60"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =570
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =885
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =2
                            Left =8340
                            Top =180
                            Width =2040
                            Height =300
                            Name ="Label22"
                            Caption ="Minutes between polls::"
                            LayoutCachedLeft =8340
                            LayoutCachedTop =180
                            LayoutCachedWidth =10380
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =15485
                    Height =1500
                    Name ="BoxScreenTop"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =15545
                    LayoutCachedHeight =1560
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10560
                    Top =570
                    Height =315
                    ColumnOrder =7
                    TabIndex =3
                    Name ="txtPollCount"
                    DefaultValue ="0"

                    LayoutCachedLeft =10560
                    LayoutCachedTop =570
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =885
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =2
                            Left =10560
                            Top =180
                            Width =1455
                            Height =300
                            Name ="Label29"
                            Caption ="Number of Polls"
                            LayoutCachedLeft =10560
                            LayoutCachedTop =180
                            LayoutCachedWidth =12015
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =1080
                    Width =1320
                    Height =315
                    ColumnOrder =2
                    TabIndex =6
                    Name ="ClubID"
                    ControlSource ="ClubID"
                    Format ="General Number"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =1395
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1560
                    Top =1080
                    Width =4335
                    Height =315
                    ColumnOrder =1
                    TabIndex =5
                    Name ="ClubName"
                    ControlSource ="ClubName"
                    Format ="@"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =1395
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6240
                    Top =1080
                    Width =1140
                    Height =315
                    ColumnOrder =0
                    TabIndex =4
                    Name ="txtOrderNumber"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =1080
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1395
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4200
                    Top =525
                    TabIndex =7
                    Name ="cmdPause"
                    Caption ="Pause"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =525
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =885
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1080
                    Width =4335
                    Height =315
                    ColumnOrder =3
                    TabIndex =8
                    Name ="txtFileName"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1080
                    LayoutCachedWidth =12015
                    LayoutCachedHeight =1395
                End
            End
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            BackColor =13233915
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =8760
            BackColor =13233915
            Name ="Detail"
            AlternateBackColor =13233915
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =105
                    Top =180
                    Width =15360
                    Height =8415
                    Name ="tabStatusForms"

                    LayoutCachedLeft =105
                    LayoutCachedTop =180
                    LayoutCachedWidth =15465
                    LayoutCachedHeight =8595
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =240
                            Top =660
                            Width =15095
                            Height =7800
                            BorderColor =10921638
                            Name ="Page49"
                            Caption ="Orders"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =660
                            LayoutCachedWidth =15335
                            LayoutCachedHeight =8460
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =300
                                    Top =720
                                    Width =15035
                                    Height =7740
                                    Name ="WebOrderStatus"
                                    SourceObject ="Form.WebOrderStatus"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =15335
                                    LayoutCachedHeight =8460
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =240
                            Top =660
                            Width =15090
                            Height =7800
                            BorderColor =10921638
                            Name ="Page33"
                            Caption ="Order guides"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =660
                            LayoutCachedWidth =15330
                            LayoutCachedHeight =8460
                            Begin
                                Begin ComboBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =5
                                    ListWidth =7200
                                    Left =1260
                                    Top =780
                                    Width =12180
                                    Height =315
                                    Name ="cmbLicensee"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Club.ID, Club.Name, Club.Contact, Club.Phone1, Club.EMail FROM Club WHERE"
                                        " (((Club.EMail) Like \"*@*\")) ORDER BY Club.[Name];"
                                    ColumnWidths ="0;2880;2880;2880;2880"
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =780
                                    LayoutCachedWidth =13440
                                    LayoutCachedHeight =1095
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =240
                                            Top =780
                                            Width =915
                                            Height =305
                                            Name ="Licensee:_Label"
                                            Caption ="Licensee:"
                                            EventProcPrefix ="Licensee__Label"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =780
                                            LayoutCachedWidth =1155
                                            LayoutCachedHeight =1085
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =11760
                                    Top =1200
                                    Width =1680
                                    Height =1440
                                    TabIndex =1
                                    Name ="cmdOGExport"
                                    Caption ="Replace Online Order Guide Based on Sales History"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =11760
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =13440
                                    LayoutCachedHeight =2640
                                    Overlaps =1
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1440
                                    Top =1260
                                    Width =10260
                                    Height =7080
                                    TabIndex =2
                                    Name ="ClubOrderGuidesWeb"
                                    SourceObject ="Form.ClubOrderGuidesWeb"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =1260
                                    LayoutCachedWidth =11700
                                    LayoutCachedHeight =8340
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =240
                            Top =660
                            Width =15090
                            Height =7800
                            BorderColor =10921638
                            Name ="Other"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =660
                            LayoutCachedWidth =15330
                            LayoutCachedHeight =8460
                            Begin
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =3900
                                    Top =840
                                    Width =2580
                                    Height =720
                                    Name ="cmdCheckNow"
                                    Caption ="Check Now!"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =3900
                                    LayoutCachedTop =840
                                    LayoutCachedWidth =6480
                                    LayoutCachedHeight =1560
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =3900
                                    Top =1740
                                    Width =2640
                                    Height =1320
                                    TabIndex =1
                                    Name ="cmdCreateTextFiles"
                                    Caption ="Refresh Item, Sales department, Size and Category"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =3900
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =6540
                                    LayoutCachedHeight =3060
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =15485
                    Height =8700
                    Name ="BoxScreen"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =15545
                    LayoutCachedHeight =8760
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            BackColor =13233915
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            Height =0
            BackColor =13233915
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "ClubOrderWebImport.cls"
