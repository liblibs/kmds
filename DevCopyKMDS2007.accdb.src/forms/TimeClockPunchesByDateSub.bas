Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =10
    ItemSuffix =35
    Left =675
    Top =1755
    Right =10965
    Bottom =5580
    RecSrcDt = Begin
        0xfb9e7e9d89dce240
    End
    RecordSource ="qryTimeClockPunchesByDateSub"
    Caption ="TimeClockPunches"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
        Begin CustomControl
            SpecialEffect =2
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
            Visible = NotDefault
            Height =255
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6000
                    Width =900
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =0
                    Name ="CashierID"
                    ControlSource ="CashierID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Width =1560
                            Height =255
                            Name ="CashierID_Label"
                            Caption ="CashierID"
                        End
                    End
                End
            End
        End
        Begin Section
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4440
                    Top =120
                    Width =900
                    ColumnWidth =840
                    ColumnOrder =2
                    TabIndex =1
                    Name ="Punch1"
                    ControlSource ="Punch1"
                    Format ="Short Time"
                    InputMask ="00:00;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =120
                            Width =585
                            Height =240
                            Name ="Label16"
                            Caption ="First In:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4380
                    Top =480
                    Width =960
                    ColumnWidth =1005
                    ColumnOrder =3
                    TabIndex =2
                    Name ="Punch2"
                    ControlSource ="Punch2"
                    Format ="Short Time"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="00:00;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =480
                            Width =705
                            Height =240
                            Name ="Label20"
                            Caption ="First Out:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6780
                    Top =120
                    ColumnWidth =1425
                    ColumnOrder =6
                    TabIndex =6
                    BackColor =10092543
                    Name ="HoursWorked"
                    ControlSource ="HoursWorked"
                    Format ="Fixed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =120
                            Width =1110
                            Height =240
                            Name ="Label21"
                            Caption ="HoursWorked:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =480
                    ColumnWidth =855
                    ColumnOrder =7
                    FontSize =1
                    TabIndex =7
                    BackColor =10092543
                    ForeColor =-2147483643
                    Name ="Wages"
                    ControlSource ="Wages"
                    Format ="$#,##0.00;($#,##0.00)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5940
                            Top =480
                            Width =630
                            Height =240
                            Name ="Label28"
                            Caption ="Wages:"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1800
                    Left =600
                    Top =120
                    Width =1020
                    ColumnWidth =645
                    ColumnOrder =0
                    Name ="Combo24"
                    ControlSource ="CashierID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Cashier].[ID], [Cashier].[Name] FROM Cashier ORDER BY [ID]; "
                    ColumnWidths ="360;1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =285
                            Height =245
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1800
                    Left =2340
                    Top =120
                    Width =1020
                    ColumnWidth =1290
                    ColumnOrder =1
                    TabIndex =5
                    BoundColumn =1
                    Name ="Combo26"
                    ControlSource ="CashierID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Cashier.Name, Cashier.ID FROM Cashier ORDER BY Cashier.Name; "
                    ColumnWidths ="360;1440"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1800
                            Top =120
                            Width =540
                            Height =240
                            Name ="Label27"
                            Caption ="Name:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =480
                    Width =2310
                    Height =255
                    ColumnWidth =1140
                    ColumnOrder =8
                    TabIndex =8
                    Name ="PunchDate"
                    ControlSource ="PunchDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =480
                            Width =1560
                            Height =255
                            Name ="PunchDate_Label"
                            Caption ="PunchDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4440
                    Top =840
                    Width =900
                    ColumnWidth =825
                    ColumnOrder =4
                    TabIndex =3
                    Name ="Punch3"
                    ControlSource ="Punch3"
                    Format ="Short Time"
                    InputMask ="00:00;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =840
                            Width =600
                            Height =240
                            Name ="Label32"
                            Caption ="Last In:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4380
                    Top =1200
                    Width =960
                    ColumnWidth =990
                    ColumnOrder =5
                    TabIndex =4
                    Name ="Punch4"
                    ControlSource ="Punch4"
                    Format ="Short Time"
                    InputMask ="00:00;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3960
                            Top =1200
                            Width =720
                            Height =240
                            Name ="Label34"
                            Caption ="Last Out:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =840
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =2100
                    ColumnWidth =1035
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3900
                            Top =60
                            Width =1560
                            Height =255
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3840
                    Top =300
                    Width =2100
                    ColumnWidth =1035
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2220
                            Top =300
                            Width =1560
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    TabIndex =2
                    Name ="SumOfHoursWorked"
                    ControlSource ="=Sum([HoursWorked])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1110
                            Height =240
                            Name ="Label23"
                            Caption ="HoursWorked:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =600
                    TabIndex =3
                    Name ="SumOfWages"
                    ControlSource ="=Sum([Wages])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2520
                            Top =600
                            Width =1110
                            Height =240
                            Name ="Label30"
                            Caption ="Wages:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TimeClockPunchesByDateSub.cls"
