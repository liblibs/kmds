Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12075
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =5505
    Top =3870
    Right =17910
    Bottom =9270
    HelpContextId =7100
    RecSrcDt = Begin
        0xb8e599438889e540
    End
    RecordSource ="Sysfile"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =5415
            BackColor =-2147483633
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =360
                    Top =240
                    Width =1485
                    Height =405
                    Name ="cmdBackup"
                    Caption ="Backup Database"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =240
                    LayoutCachedWidth =1845
                    LayoutCachedHeight =645
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =360
                    Top =1080
                    Width =1485
                    Height =405
                    TabIndex =1
                    Name ="cmdRestore"
                    Caption ="Restore Database"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1845
                    LayoutCachedHeight =1485
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =10440
                    Top =3240
                    Width =666
                    Height =606
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdfrmClose"
                    Caption ="Command64"
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
                        0x0707070707070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109

                    LayoutCachedLeft =10440
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11106
                    LayoutCachedHeight =3846
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =10440
                            Top =3840
                            Width =885
                            Height =285
                            Name ="Label78"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10440
                            LayoutCachedTop =3840
                            LayoutCachedWidth =11325
                            LayoutCachedHeight =4125
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =3300
                    Width =3660
                    Height =585
                    TabIndex =3
                    Name ="cmdOnlineBackup"
                    Caption ="Online Backup (requires an internet connection)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =3885
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =3960
                    Width =3660
                    Height =585
                    TabIndex =4
                    Name ="Command10"
                    Caption ="Verify Online Backup (after Online Backup)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =4545
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2460
                    Top =360
                    Width =9360
                    Height =2700
                    TabIndex =5
                    Name ="KMDSLinks"
                    SourceObject ="Form.KMDSLinks"

                    LayoutCachedLeft =2460
                    LayoutCachedTop =360
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2520
                            Width =915
                            Height =240
                            Name ="Label11"
                            Caption ="Useful Links"
                            LayoutCachedLeft =2520
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4980
                    Top =3300
                    Width =3660
                    Height =585
                    TabIndex =6
                    Name ="Command12"
                    Caption ="Get KMDS Current Version "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =3300
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =3885
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =4080
                    Width =3720
                    Height =480
                    TabIndex =7
                    Name ="cmdChronJob"
                    Caption ="Create Scheduled Task for ftp.kmds.us backup"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =4080
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =4560
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8760
                    Top =4140
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtAgentNumber"
                    ControlSource ="AgentNumber"
                    GridlineColor =10921638

                    LayoutCachedLeft =8760
                    LayoutCachedTop =4140
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =4455
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8760
                    Top =4560
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtStartTime"
                    Format ="Short Time"
                    GridlineColor =10921638

                    LayoutCachedLeft =8760
                    LayoutCachedTop =4560
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =4875
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7320
                            Top =4620
                            Width =1455
                            Height =210
                            Name ="Label22"
                            Caption ="Agent time off-set:"
                            LayoutCachedLeft =7320
                            LayoutCachedTop =4620
                            LayoutCachedWidth =8775
                            LayoutCachedHeight =4830
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =600
                    Top =4920
                    Width =9720
                    Height =360
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSCHTask"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =4920
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =5280
                End
            End
        End
    End
End
CodeBehindForm
' See "BackupSQLDatabase.cls"
