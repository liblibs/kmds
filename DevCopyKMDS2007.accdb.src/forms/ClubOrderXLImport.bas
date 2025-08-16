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
    Width =9060
    DatasheetFontHeight =11
    ItemSuffix =16
    Top =600
    Right =10845
    Bottom =7095
    RecSrcDt = Begin
        0xee1696443270e340
    End
    RecordSource ="qryClubOrderXL"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =5137573
            LabelX =-1800
        End
        Begin Section
            Height =4440
            BackColor =13233915
            Name ="Detail"
            AlternateBackColor =13233915
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =540
                    Width =3480
                    Height =3015
                    Name ="Text0"

                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =3555
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =120
                            Width =1770
                            Height =315
                            Name ="Label1"
                            Caption ="Spreadsheets in "
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =1890
                            LayoutCachedHeight =435
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =240
                    Top =3840
                    Width =3720
                    Height =390
                    TabIndex =1
                    Name ="Command2"
                    Caption ="Create dispenser orders"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =4230
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =5940
                    Top =780
                    Width =2940
                    Height =315
                    TabIndex =2
                    Name ="txtFileName"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =780
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =1095
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =3600
                            Top =780
                            Width =2370
                            Height =315
                            Name ="Label4"
                            Caption ="Currently Processing..."
                            LayoutCachedLeft =3600
                            LayoutCachedTop =780
                            LayoutCachedWidth =5970
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5880
                    Top =1260
                    Width =2940
                    Height =315
                    TabIndex =3
                    Name ="ClubID"
                    ControlSource ="ClubID"
                    Format ="General Number"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =1260
                            Width =1320
                            Height =315
                            Name ="Label10"
                            Caption ="Club ID"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =1260
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =1575
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5820
                    Top =1920
                    Width =2940
                    Height =315
                    TabIndex =4
                    Name ="ClubName"
                    ControlSource ="ClubName"
                    Format ="@"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1920
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =2235
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =1920
                            Width =1320
                            Height =315
                            Name ="Label11"
                            Caption ="ClubName"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =2235
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5820
                    Top =2580
                    Width =2940
                    Height =315
                    TabIndex =5
                    Name ="txtOrderNumber"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =2580
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4440
                            Top =2580
                            Width =1320
                            Height =315
                            Name ="Label13"
                            Caption ="Order"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =2580
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =2895
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =60
                    Top =60
                    Width =9000
                    Height =4380
                    Name ="BoxScreen"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =4440
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrderXLImport.cls"
