Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    ScrollBars =2
    ViewsAllowed =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4140
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =7245
    Top =2505
    Right =13095
    Bottom =10275
    DatasheetBackColor =16104334
    OrderBy ="[AgentsActiveForTransferFalse].[LastName], [AgentsActiveForTransferFalse].[Agent"
        "Number]"
    RecSrcDt = Begin
        0x3632dd704c6ee340
    End
    RecordSource ="SELECT Agents.AgentNumber, Agents.FirstName, Agents.LastName, Agents.ActiveForTr"
        "ansfer FROM Agents WHERE ((Not (Agents.ActiveForTransfer)=Yes)); "
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowFormView =0
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483603
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =-2147483603
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin Section
            Height =4980
            BackColor =-2147483603
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
                    ColumnWidth =1725
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =600
                            Width =1455
                            Height =315
                            Name ="Label0"
                            Caption ="Agent #:"
                            LayoutCachedLeft =300
                            LayoutCachedTop =600
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1020
                    Height =315
                    ColumnWidth =2580
                    TabIndex =1
                    Name ="LastName"
                    ControlSource ="LastName"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1020
                            Width =1170
                            Height =315
                            Name ="Label1"
                            Caption ="Store Name"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1440
                    Height =315
                    ColumnWidth =1770
                    TabIndex =2
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1440
                            Width =1095
                            Height =315
                            Name ="Label2"
                            Caption ="Agent"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =1755
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AgentsActiveForTransferFalse.cls"
