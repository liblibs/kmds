Attribute VB_Name = "modAgeCalc"
Option Compare Database
Option Explicit

Public Function Age(Dob As Date, today As Date, Optional WithMonths As Boolean = False, _
             Optional WithDays As Boolean = False, Optional DisplayWithWords As Boolean = False) As Variant
   'Author: © Copyright 2001 Pacific Database Pty Limited
   ' Graham R Seach gseach@pacificdb.com.au
   ' Phone: +61 2 9872 9594 Fax: +61 2 9872 9593
   '
   ' You may freely use and distribute this code
   ' with any applications you may develop, on the
   ' condition that the copyright notice remains
   ' unchanged, and intact as part of the code. You
   ' may not sell or publish this code in any form
   ' without the express written permission of the
   ' copyright holder.
   '
   ' Description: This function calculates a person's age,
   ' given their date of birth, and a second date.
   '
   ' Inputs: DOB: The person's date of birth
   ' Today: The second date (ostensibly today)
   ' WithMonths: Boolean - If True, displays months
   ' DisplayWithWords: Boolean - If True, displays
   ' (ie: years / months)
   '
   ' Outputs: On error: Null
   ' On no error: Variant containing person's age in
   ' years, months and days (if selected).
   ' If DisplayWithWords = False:
   ' Months and days, if selected, are shown
   ' to the right of the decimal point, but
   ' are the actual number of months and days,
   ' not a fraction of the year. For example,
   ' 44.11.03 = 44 years 11 months and 3 days.
   ' If DisplayWithWords = True:
   ' Output example: "44 years 11 months 3 days",
   ' except where months = 0, in which case, no
   ' months are shown.

   If gcfHandleErrors Then On Error GoTo Age_ErrorHandler

   Dim iYears As Integer
   Dim iMonths As Integer
   Dim iDays As Integer
   Dim dTempDate As Date

   ' Check that the dates are valid
   If Not (IsDate(Dob)) Or Not (IsDate(today)) Then
      DoCmd.Beep
      MsgBox "Invalid date.", vbOKOnly + vbInformation, "Invalid date"
      Exit Function
   End If

   ' Check that DOB < Today
   If Dob > today Then
      DoCmd.Beep
      MsgBox "Today must be greater than DOB.", _
         vbOKOnly + vbInformation, "Invalid date position"
      GoTo Age_ErrorHandler
   End If

   iYears = Abs(DateDiff("yyyy", Dob, today) - _
      IIf(Format(Dob, "mmdd") <= Format(today, "mmdd"), 0, 1))
   dTempDate = DateAdd("yyyy", iYears, Dob)

   If WithMonths Then
      iMonths = DateDiff("m", dTempDate, today) - _
         IIf(DateAdd("m", iMonths, DateAdd("yyyy", iYears, Dob)) > today, 1, 0)
      dTempDate = DateAdd("m", iMonths, dTempDate)
   End If

   If WithDays Then
      iDays = today - dTempDate
   End If

   ' Format the output
   If DisplayWithWords Then
      ' Display the output in words
      Age = IIf(iYears > 0, iYears & " year" & IIf(iYears <> 1, "s", ""), "")
      Age = Age & IIf(WithMonths, iMonths & " month" & IIf(iMonths <> 1, "s", ""), "")
      Age = Trim(Age & IIf(WithDays, iDays & " day" & IIf(iDays <> 1, "s", ""), ""))
   Else
      ' Display the output in the format yy.mm.dd
      Age = Trim(iYears & IIf(WithMonths, "." & Format(iMonths, "00"), "") _
      & IIf(WithDays, "." & Format(iDays, "00"), ""))
   End If

Exit_Age:
      Exit Function

Age_ErrorHandler:
      Age = Null
End Function

Public Function Diff2Dates(Interval As String, Date1 As Date, Date2 As Date, _
    Optional ShowZero As Boolean = False) As Variant
    ' Author:    ©Copyright 2001 Pacific Database Pty Limited
    '           Graham R Seach MCP MVP gseach@pacificdb.com.au
    '           Phone: +61 2 9872 9594  Fax: +61 2 9872 9593
    '           (*) Amendments suggested by Douglas J. Steele MVP
    '
    '           You may freely use and distribute this code
    '           with any applications you may develop, on the
    '           condition that the copyright notice remains
    '           unchanged, and intact as part of the code. You
    '           may not sell or publish this code in any form
    '           without the express written permission of the
    '           copyright holder.
    '
    ' Description:   This function calculates the number of years,
    '           months, days, hours, minutes and seconds between
    '           two dates, as elapsed time.
    '
    ' Inputs:    Interval:   Intervals to be displayed (a string)
    '           Date1:      The lower date (see below)
    '           Date2:      The higher date (see below)
    '           ShowZero:   Boolean to select showing zero elements
    '
    ' Outputs:   On error: Null
    '           On no error: Variant containing the number of years,
    '               months, days, hours, minutes & seconds between
    '               the two dates, depending on the display interval
    '               selected.
    '           If Date1 is greater than Date2, the result will
    '               be a negative value.
    '           The function compensates for the lack of any intervals
    '               not listed. For example, if Interval lists "m", but
    '               not "y", the function adds the value of the year
    '               component to the month component.
    '           If ShowZero is True, and an output element is zero, it
    '               is displayed. However, if ShowZero is False or
    '               omitted, no zero-value elements are displayed.
    '               For example, with ShowZero = False, Interval="ym",
    '               elements = 0 & 1 respectively, the output string
    '               will be "1 month" - not "0 years 1 month".

    ' Additional changes:
    ' 1) Formats associated with date segments
    ' 2) Changed order of arguments

    If gcfHandleErrors Then On Error GoTo Err

    Dim varTemp As Variant, baseDate As Date
    Dim diffY As Long, diffM As Long, diffD As Long
    Dim diffH As Long, diffN As Long, diffS As Long
    Dim Y As Boolean, m As Boolean, d As Boolean
    Dim h As Boolean, n As Boolean, S As Boolean
    Dim ctr As Integer, tmpDate As Date, swapped As Boolean

    '***********************************************
    'Change the following constants according to the
    'desired output language
    Const Yvar = " year": Const Yvars = " years"
    Const Mvar = " month": Const Mvars = " months"
    Const Dvar = " day": Const Dvars = " days"
    Const Hvar = " hour": Const Hvars = " hours"
    Const Nvar = " minute": Const Nvars = " minutes"
    Const Svar = " second": Const Svars = " seconds"

    Const INTERVALS As String = "dmyhns"
    'Check that Interval contains valid characters
    For ctr = 1 To Len(Interval)
        If InStr(1, INTERVALS, Mid(Interval, ctr, 1)) = 0 Then
            Exit Function
        End If
    Next ctr

    'Check that valid dates have been entered
    If Not (IsDate(Date1)) Then Exit Function
    If Not (IsDate(Date2)) Then Exit Function

    'If necessary, swap the dates, to ensure that
    'Date1 is lower than Date2
    If Date1 > Date2 Then
        tmpDate = Date1
        Date1 = Date2
        Date2 = tmpDate
        swapped = True
    End If

    Diff2Dates = Null
    varTemp = Null

    'What intervals are supplied
    Y = (InStr(1, Interval, "y") > 0)
    m = (InStr(1, Interval, "m") > 0)
    d = (InStr(1, Interval, "d") > 0)
    h = (InStr(1, Interval, "h") > 0)
    n = (InStr(1, Interval, "n") > 0)
    S = (InStr(1, Interval, "s") > 0)

    'If not gcfHandleErrors then Debug.Print "Date1: " & Date1
    'If not gcfHandleErrors then Debug.Print "Date2: " & Date2

    'Get the cumulative differences
    If Y Then
        diffY = Abs(DateDiff("yyyy", Date1, Date2)) - _
                IIf(Format(Date1, "mmddhhnnss") <= Format(Date2, "mmddhhnnss"), 0, 1) '**
        Date1 = DateAdd("yyyy", diffY, Date1)
    End If
    If m Then
        diffM = Abs(DateDiff("m", Date1, Date2)) - _
                IIf(Format(Date1, "ddhhnnss") <= Format(Date2, "ddhhnnss"), 0, 1) '**
        Date1 = DateAdd("m", diffM, Date1)
    End If
    If d Then
        diffD = Abs(DateDiff("d", Date1, Date2)) - _
                IIf(Format(Date1, "hhnnss") <= Format(Date2, "hhnnss"), 0, 1) '**
        Date1 = DateAdd("d", diffD, Date1)
    End If
    If h Then
        diffH = Abs(DateDiff("h", Date1, Date2)) - _
                IIf(Format(Date1, "nnss") <= Format(Date2, "nnss"), 0, 1) '**
        Date1 = DateAdd("h", diffH, Date1)
    End If
    If n Then
        diffN = Abs(DateDiff("n", Date1, Date2)) - _
                IIf(Format(Date1, "ss") <= Format(Date2, "ss"), 0, 1) '**
        Date1 = DateAdd("n", diffN, Date1)
    End If
    If S Then
        diffS = Abs(DateDiff("s", Date1, Date2))
        Date1 = DateAdd("s", diffS, Date1)
    End If

    'Set the output display
    If Y And (diffY > 0 Or ShowZero) Then
        varTemp = IIf(swapped, IIf(diffY > 0, -diffY, diffY), diffY) & _
                    IIf(diffY <> 1, Yvars, Yvar)
    End If
    If m And (diffM > 0 Or ShowZero) Then
        If m Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & _
                        diffM & IIf(diffM <> 1, Mvars, Mvar)
        End If
    End If
    If d And (diffD > 0 Or ShowZero) Then
        If d Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & _
                        diffD & IIf(diffD <> 1, Dvars, Dvar)
        End If
    End If
    If h And (diffH > 0 Or ShowZero) Then
        If h Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & _
                        diffH & IIf(diffH <> 1, Hvars, Hvar)
        End If
    End If
    If n And (diffN > 0 Or ShowZero) Then
        If n Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & _
                        diffN & IIf(diffN <> 1, Nvars, Nvar)
        End If
    End If
    If S And (diffS > 0 Or ShowZero) Then
        If S Then
            varTemp = varTemp & IIf(IsNull(varTemp), Null, " ") & _
                        diffS & IIf(diffS <> 1, Svars, Svar)
        End If
    End If

    Diff2Dates = Trim(varTemp)
    Exit Function
Err:
    Diff2Dates = Null
End Function
