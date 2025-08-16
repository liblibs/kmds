Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =2130
    Top =2250
    Right =11415
    Bottom =5955
    RecSrcDt = Begin
        0xc27a2c1cecc1e340
    End
    RecordSource ="SELECT Sysfile.OLCCFTPServer, Sysfile.OLCCFTPUser, Sysfile.OLCCFTPPassword, Sysf"
        "ile.OLCCFTPEndOfDaySend, Sysfile.OLCCFTPASCII FROM Sysfile; "
    DatasheetFontName ="Calibri"
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
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin Section
            Height =3571
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2745
                    Top =360
                    Width =3600
                    Height =315
                    Name ="OLCCFTPServer"
                    ControlSource ="OLCCFTPServer"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2745
                    LayoutCachedTop =360
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =2325
                            Height =315
                            Name ="Label0"
                            Caption ="OLCCFTPServer:"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =675
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2745
                    Top =750
                    Width =3600
                    Height =315
                    TabIndex =1
                    Name ="OLCCFTPUser"
                    ControlSource ="OLCCFTPUser"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2745
                    LayoutCachedTop =750
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =1065
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =750
                            Width =2325
                            Height =315
                            Name ="Label1"
                            Caption ="OLCCFTPUser:"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =750
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =1065
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2745
                    Top =1140
                    Width =3600
                    Height =315
                    TabIndex =2
                    Name ="OLCCFTPPassword"
                    ControlSource ="OLCCFTPPassword"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2745
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =1455
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1140
                            Width =2325
                            Height =315
                            Name ="Label2"
                            Caption ="OLCCFTPPassword:"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =1455
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2745
                    Top =1530
                    Width =3600
                    Height =315
                    TabIndex =3
                    Name ="OLCCFTPEndOfDaySend"
                    ControlSource ="OLCCFTPEndOfDaySend"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2745
                    LayoutCachedTop =1530
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =1845
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1530
                            Width =2325
                            Height =315
                            Name ="Label3"
                            Caption ="OLCCFTPEndOfDaySend:"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1530
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =1845
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2745
                    Top =1920
                    Width =3600
                    Height =315
                    TabIndex =4
                    Name ="OLCCFTPASCII"
                    ControlSource ="OLCCFTPASCII"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2745
                    LayoutCachedTop =1920
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =2235
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1920
                            Width =2325
                            Height =315
                            Name ="Label4"
                            Caption ="OLCCFTPASCII:"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1920
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =2235
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3344
                    Top =2494
                    Width =576
                    Height =576
                    TabIndex =5
                    Name ="Command5"
                    Caption ="Command5"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b06040d0905830c090583030 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b0704030a05840ff905030ff ,
                        0x905030e09058305000000000000000000000000000000000905830ff00000000 ,
                        0x905830ff00000000a05840ffb06040ffc06840ffc07050ffd07850ffb06040ff ,
                        0x905830ff905030ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000f0886040e08050ffe07840ff ,
                        0xc0704090b0684020000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000e08850ffe08860ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000a07050ff703010ff703010ff703010ff702800ff ,
                        0x00000000702800ff602800ff602000ff602000ff602000ff0000000000000000 ,
                        0x000000000000000000000000a07860fff0d8d0fffffffffff0d0b0ffa06850ff ,
                        0x00000000906040ffe0b8a0ffe0d0c0ffd0a080ff602000ff0000000000000000 ,
                        0x000000000000000000000000b08060fff0d0c0fffffffffff0d0b0ffa07050ff ,
                        0x00000000a06850ffe0c8b0ffe0d8d0ffd09880ff704020ff0000000000000000 ,
                        0x000000000000000000000000c09880ffb08070ffb08870ffa07860ffa07860ff ,
                        0x703010ffa06850ffc09880ffa08060ff906850ff805840ff0000000000000000 ,
                        0x00000000000000000000000000000000b09080fffffffffff0d0b0ff906040ff ,
                        0x703010ff906040fffff8f0fff0c0b0ff805030ff000000000000000000000000 ,
                        0x00000000000000000000000000000000c09880fffffffffff0d0b0ffb07860ff ,
                        0x804820ff906040fffffffffff0c8b0ff805030ff000000000000000000000000 ,
                        0x00000000000000000000000000000000c0a090ffc09880ffb08060ffc09080ff ,
                        0xa07050ffb08870ffc09880ffa07860ff906850ff000000000000000000000000 ,
                        0x00000000000000000000000000000000c0a09060c0a090ffe0d8d0ffc09080ff ,
                        0xc09080ffc09880fff0f0f0ff906040ffa0685040000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000d0a890ffffffffffd0a890ff ,
                        0x00000000d0a890ffffffffff905840ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000c09880ffb09070ffb08870ff ,
                        0x00000000c09880ffb09070ffb08870ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3344
                    LayoutCachedTop =2494
                    LayoutCachedWidth =3920
                    LayoutCachedHeight =3070
                End
            End
        End
    End
End
CodeBehindForm
' See "Form4.cls"
