Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14820
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =-3600
    Top =735
    Right =13005
    Bottom =9360
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x22c32a4eb3dee240
    End
    RecordSource ="qryLotteryScratchRecon"
    Caption ="Lottery Scratch Recon"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =7680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =247
                    Left =360
                    Top =780
                    Width =13860
                    Height =6600
                    Name ="LotteryScratch"
                    SourceObject ="Form.LotteryScratchSub"
                    LinkChildFields ="BusinessDate"
                    LinkMasterFields ="BusinessDate"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =540
                            Width =1110
                            Height =240
                            Name ="Label2"
                            Caption ="LotteryScratch"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =1860
                    Top =240
                    Width =2400
                    Height =420
                    FontSize =16
                    TabIndex =1
                    Name ="cmbBusinessDate"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [qryLotteryScratchRecon].[BusinessDate] FROM qryLotteryScratchRecon; "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =420
                            Top =240
                            Width =1185
                            Height =245
                            Name ="BusinessDate1_Label"
                            Caption ="BusinessDate"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6420
                    Top =300
                    TabIndex =2
                    Name ="BusinessDate"
                    ControlSource ="BusinessDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4980
                            Top =300
                            Width =1095
                            Height =240
                            Name ="Label13"
                            Caption ="BusinessDate:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =14280
                    Height =7440
                    Name ="Box22"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10020
                    Top =240
                    Width =1800
                    Height =420
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    BackColor =16777164
                    Name ="Text24"
                    ControlSource ="=LotteryScratch.Form!TotSoldDollars"
                    Format ="Standard"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8340
                            Top =240
                            Width =1635
                            Height =240
                            Name ="Label25"
                            Caption ="Cash register receipts:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =696
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1320
                    Top =420
                    FontWeight =700
                    Name ="Text3"
                    ControlSource ="=LotteryScratch.Form!TotBeginMeter"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =120
                            Width =645
                            Height =240
                            Name ="Label4"
                            Caption ="Begin"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =420
                    FontWeight =700
                    TabIndex =1
                    Name ="Text5"
                    ControlSource ="=LotteryScratch.Form!TotEndMeter"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =120
                            Width =645
                            Height =240
                            Name ="Label6"
                            Caption ="End.."
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4320
                    Top =420
                    FontWeight =700
                    TabIndex =2
                    Name ="Text7"
                    ControlSource ="=LotteryScratch.Form!TotAddMeter"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4320
                            Top =120
                            Width =645
                            Height =240
                            Name ="Label8"
                            Caption ="Added."
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7320
                    Top =420
                    FontWeight =700
                    TabIndex =3
                    Name ="Text9"
                    ControlSource ="=LotteryScratch.Form!TotAmountSold"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7320
                            Top =120
                            Width =750
                            Height =240
                            Name ="Label10"
                            Caption ="Register.."
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5820
                    Top =420
                    FontWeight =700
                    TabIndex =4
                    Name ="Text14"
                    ControlSource ="=LotteryScratch.Form!TotPhysicalAmount"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5820
                            Top =120
                            Width =705
                            Height =240
                            Name ="Label15"
                            Caption ="Count"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8940
                    Top =420
                    FontWeight =700
                    TabIndex =5
                    BackColor =16777164
                    Name ="Text16"
                    ControlSource ="=LotteryScratch.Form!TotVariance"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8940
                            Top =120
                            Width =705
                            Height =240
                            Name ="Label17"
                            Caption ="Variance"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10500
                    Top =420
                    FontWeight =700
                    TabIndex =6
                    BackColor =16777164
                    Name ="Text18"
                    ControlSource ="=LotteryScratch.Form!TotVarianceDollars"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10500
                            Top =120
                            Width =1110
                            Height =240
                            Name ="Label19"
                            Caption ="Variance $$$.."
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12120
                    Top =120
                    Width =726
                    Height =576
                    TabIndex =7
                    Name ="Command20"
                    Caption ="Print Report"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =120
                    Width =1020
                    Height =540
                    TabIndex =8
                    Name ="cmdSetupNew"
                    Caption ="Setup New Game"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13800
                    Top =60
                    Width =576
                    Height =576
                    TabIndex =9
                    Name ="Command23"
                    Caption ="Command23"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"

                End
            End
        End
    End
End
CodeBehindForm
' See "LotteryScratch.cls"
