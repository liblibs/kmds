Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    ScrollBars =2
    ViewsAllowed =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =825
    Top =2505
    Right =6675
    Bottom =10275
    DatasheetBackColor =62207
    RecSrcDt = Begin
        0x119e59694c6ee340
    End
    RecordSource ="SELECT Agents.AgentNumber, Agents.FirstName, Agents.LastName, Agents.ActiveForTr"
        "ansfer FROM Agents WHERE (((Agents.ActiveForTransfer)=Yes)); "
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowFormView =0
    FilterOnLoad =0
    DatasheetBackColor12 =62207
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =62207
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
            Height =4860
            BackColor =62207
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
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
                    ColumnWidth =2265
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
                            Caption ="Agent:"
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
' See "AgentsActiveForTransferTrue.cls"
