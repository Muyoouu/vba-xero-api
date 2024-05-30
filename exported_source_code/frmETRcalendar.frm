VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmETRcalendar 
   Caption         =   "frmETRcalendar"
   ClientHeight    =   5895
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "frmETRcalendar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmETRcalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Credits: Siddharth Rout, logicworkz
' Source: https://stackoverflow.com/questions/54650417/how-can-i-create-a-calendar-input-in-vba-excel


    Option Explicit
    
    Private TimerID As Long, TimerSeconds As Single, tim As Boolean
    Dim curDate As Date
    Dim i As Long
    Dim thisDay As Integer, thisMonth As Integer, thisYear As Integer
    Dim CBArray() As New CalendarClass
        
    '*JCR
        Dim Darray() As New CalendarClass
    '**********
        
    Dim NewXpos As Single
    Dim NewYpos As Single

    Private Cal_theme As CalendarThemes
    Private LdtFormat As String, SdtFormat As String
    Private pUserSelectedDateStr As String
    Private pSelectedCalTheme As Byte
    
    Public Property Let LongDateFormat(s As String)
        LdtFormat = s
        
        lblTitleCurDt.Caption = Format(Date, LdtFormat)
    End Property
    
    Public Property Get LongDateFormat() As String
        LongDateFormat = LdtFormat
    End Property
    
    Public Property Let ShortDateFormat(s As String)
        SdtFormat = s
    End Property
    
    Public Property Get ShortDateFormat() As String
        ShortDateFormat = SdtFormat
    End Property
    
    Public Property Let Caltheme(Theme As CalendarThemes)
        Cal_theme = Theme
        
        '--> Set the color of controls
        Select Case Cal_theme
            Case CalendarThemes.Venom
                MyBackColor = RGB(69, 69, 69)
                MyForeColor = RGB(252, 248, 248)
                CurDateColor = RGB(246, 127, 8)
                CurDateForeColor = RGB(0, 0, 0)
                NotCurDateColor = RGB(90, 90, 90)
                
            Case CalendarThemes.MartianRed
                MyBackColor = RGB(87, 0, 0)
                MyForeColor = RGB(203, 146, 146)
                CurDateColor = RGB(122, 185, 247)
                CurDateForeColor = RGB(0, 0, 0)
                NotCurDateColor = RGB(116, 0, 0)
                
            Case CalendarThemes.ArcticBlue
                MyBackColor = RGB(42, 48, 92)
                MyForeColor = RGB(179, 179, 179)
                CurDateColor = RGB(122, 185, 247)
                CurDateForeColor = RGB(0, 0, 0)
                NotCurDateColor = RGB(66, 71, 118)
                
            Case CalendarThemes.Greyscale
                MyBackColor = RGB(240, 240, 240)
                MyForeColor = RGB(0, 0, 0)
                CurDateColor = RGB(246, 127, 8)
                CurDateForeColor = RGB(0, 0, 0)
                NotCurDateColor = RGB(225, 225, 225)
                
        End Select
        
        Me.BackColor = MyBackColor
        FrameDay.BackColor = MyBackColor
        FrameMonth.BackColor = MyBackColor
        FrameYr.BackColor = MyBackColor
        
        lblTitleCurDt.ForeColor = CurDateColor
        
        lblTitleCurMY.ForeColor = MyForeColor
        lblTitleCurMY.BorderColor = MyForeColor
        
        lblTitleClock.ForeColor = MyForeColor
        lblTitleAMPM.ForeColor = MyForeColor
        lblUnload.ForeColor = MyForeColor
        lblThemes.ForeColor = MyForeColor
        
        lblUP.ForeColor = MyForeColor
        lblDOWN.ForeColor = MyForeColor
        
        '--> Days
        For i = 1 To 42
            With Me.Controls("D" & i)
                .ForeColor = MyForeColor
                .BorderColor = MyForeColor
            End With
        Next i
        
        '--> Weekdays
        For i = 1 To 7
            With Me.Controls("WD" & i)
                .ForeColor = MyForeColor
            End With
        Next i
        
        '--> Month
        For i = 1 To 12
            With Me.Controls("M" & i)
                .ForeColor = MyForeColor
                .BorderColor = MyForeColor
            End With
        Next i
        
        '--> Year
        For i = 1 To 12
            With Me.Controls("Y" & i)
                .ForeColor = MyForeColor
                .BorderColor = MyForeColor
            End With
        Next i
        
        '--> Populate this months calendar
        PopulateCalendar Date
    End Property

    Public Property Get Caltheme() As CalendarThemes
        Caltheme = Cal_theme
    End Property
    
    Public Property Get UserSelectedDateStr() As String
        UserSelectedDateStr = pUserSelectedDateStr
    End Property
    
    Public Property Let UserSelectedDateStr(s As String)
        pUserSelectedDateStr = s
    End Property



'--> allow user to cycle thru avialable themes
Private Sub lblThemes_Click()
    Dim t As Byte
    
    t = pSelectedCalTheme
    If t <= 2 Then
        t = t + 1
        frmETRcalendar.Caltheme = t
        frmETRcalendar.Repaint
        pSelectedCalTheme = t
    Else
        frmETRcalendar.Caltheme = 0
        frmETRcalendar.Repaint
        pSelectedCalTheme = 0
    End If
End Sub
'--> Unload form
Private Sub lblUnload_Click()
    Unload Me
End Sub



    Private Sub UserForm_Initialize()
        
        '--> remove borders from day labels. i keep them in place for the dev environment.
        Dim lblCtrl As control
        i = 0
        For Each lblCtrl In Me.Controls
            If TypeOf lblCtrl Is MSForms.Label Then
                lblCtrl.BorderStyle = fmBorderStyleNone
            End If
        Next
        
        '--> Hide the Title Bar
        HideTitleBar Me
        
        Me.LongDateFormat = "dddd mm, yyyy"
        Me.ShortDateFormat = "mm/dd/yyyy"
        
        '--> Create a command button control array so that
        '--> when we press escape, we can unload the userform
        Dim CBCtl As control
        
        i = 0
        
        '*JCR
        For Each CBCtl In Me.Controls
            If TypeOf CBCtl Is MSForms.Label Then
                i = i + 1
                ReDim Preserve Darray(1 To i)
                Set Darray(i).CommandButtonEvents = CBCtl
            End If
        Next CBCtl
        Set CBCtl = Nothing
        
        '***********
        
        '~~> Set the Time
        StartTimer
                  
        curDate = Date
        
        thisDay = Day(Date): thisMonth = Month(Date): thisYear = Year(Date)
         
        CurYear = Year(Date): CurMonth = Month(Date)
        
        PopulateCalendar curDate
    End Sub
    
    '--> The below 4 procedures will assist in moving the borderless userform
    Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button = 1 Then
            NewXpos = X
            NewYpos = Y
        End If
    End Sub
    Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button And 1 Then
            Me.Left = Me.Left + (X - NewXpos)
            Me.Top = Me.Top + (Y - NewYpos)
        End If
        
        lblDOWN.ForeColor = MyForeColor
        lblUP.ForeColor = MyForeColor
        lblUnload.ForeColor = MyForeColor
        lblThemes.ForeColor = MyForeColor
        
    End Sub
    Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button = 1 Then
            NewXpos = X
            NewYpos = Y
        End If
    End Sub
    
    '--> Stop timer in the terminate event
    Private Sub UserForm_Terminate()
        EndTimer
    End Sub
    
    '--> UP Button
    Private Sub lblUP_Click()
    
        Select Case Label5.Caption
            Case 1 '~~> When user presses the up button when the dates are displayed
                curDate = DateSerial(CurYear, CurMonth, 0)
                
                '~~> Check if date is >= 1/1/1919
                If curDate >= DateSerial(1919, 1, 1) Then
                    '~~> Populate prev months calendar
                    PopulateCalendar curDate
                End If
            Case 2 '<~~ Do nothing
            Case 3 '~~> When user presses the up button when the Year Range is displayed
                If frmYr > 1919 Then
                    
                    Dim NewToYr As Integer
                    
                    ToYr = frmYr - 1
                    NewToYr = frmYr - 1
                    
                    For i = 1 To 12
                        Me.Controls("Y" & i).Caption = ""
                    Next i
                
                    For i = 12 To 1 Step -1
                        If Not NewToYr < 1919 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewToYr
                                
                                If NewToYr = thisYear Then
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = CurDateColor
                                Else
                                    .BackStyle = fmBackStyleTransparent
                                End If

                                NewToYr = NewToYr - 1
                            End With
                        End If
                    Next i
                    
                    frmYr = NewToYr + 1
                    lblTitleCurMY.Caption = (NewToYr + 1) & " - " & ToYr
                End If
        End Select
    End Sub
    
    '--> Down Button
    Private Sub lblDOWN_Click()
        Select Case Label5.Caption
            Case 1 '~~> When user presses the down button when the dates are displayed
                curDate = DateAdd("m", 1, DateSerial(CurYear, CurMonth, 1))
                
                '~~> Check if date is <= 31/12/2119
                If curDate <= DateSerial(2119, 12, 31) Then
                    '~~> Populate prev months calendar
                    PopulateCalendar curDate
                End If
            Case 2 '<~~ Do nothing
            Case 3 '~~> When user presses the down button when the Year Range is displayed
                frmYr = Val(Split(lblTitleCurMY.Caption, "-")(0))
                ToYr = Val(Split(lblTitleCurMY.Caption, "-")(1))
                 
                If ToYr < 2119 Then
                    
                    Dim NewFrmYr As Integer
                    
                    frmYr = ToYr + 1
                    NewFrmYr = ToYr + 1
                    
                    For i = 1 To 12
                        Me.Controls("Y" & i).Caption = ""
                    Next i
                
                    For i = 1 To 12
                        If NewFrmYr < 2119 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewFrmYr
                                
                                If NewFrmYr = thisYear Then
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = CurDateColor
                                Else
                                    .BackStyle = fmBackStyleTransparent
                                End If
                                
                                NewFrmYr = NewFrmYr + 1
                            End With
                            
                        ElseIf NewFrmYr = 2119 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewFrmYr
                                NewFrmYr = NewFrmYr + 1
                            End With
                        End If
                    Next i
                    
                    If NewFrmYr = 2119 Then ToYr = NewFrmYr Else ToYr = NewFrmYr - 1
                    lblTitleCurMY.Caption = frmYr & " - " & ToYr
                End If
        End Select
    End Sub
    
    '--> Populate the calendar for a specific month
    Sub PopulateCalendar(d As Date)
        
        Dim m As Integer, Y As Integer
        Dim i As Integer, j As Integer
        Dim LastDay As Integer, NextCounter As Integer, PrevCounter As Integer
        Dim dtOne As Date, dtLast As Date, dtNext As Date
        
        CurYear = Year(d)
        CurMonth = Month(d)
        
        m = Month(d): Y = Year(d)
        
        '--> 1st day of the current month
        dtOne = DateSerial(Y, m, 1)
        '--> last day of the previous month
        dtLast = DateSerial(Year(dtOne), Month(dtOne), 0)
        '--> 1st day of the next month
        dtNext = DateAdd("m", 1, DateSerial(Year(dtOne), Month(dtOne), 1))
        
        '--> Set the 1st day of the month to its proper weekday
        Select Case Weekday(dtOne, 0)
            Case 1
                NextCounter = 1: PrevCounter = 0
                
            Case 2
                NextCounter = 2: PrevCounter = 1
                
            Case 3
                NextCounter = 3: PrevCounter = 2
                
            Case 4
                NextCounter = 4: PrevCounter = 3
                
            Case 5
                NextCounter = 5: PrevCounter = 4
                
            Case 6
                NextCounter = 6: PrevCounter = 5
                
            Case 7
                NextCounter = 7: PrevCounter = 6
                
        End Select
        
        '--> Get the last day of the current month
        LastDay = Val(Format(Excel.Application.WorksheetFunction.EoMonth(dtOne, 0), "dd"))
        
        '--> Populate all days for the current month
        For i = 1 To LastDay
            Me.Controls("D" & NextCounter).Caption = i
            Me.Controls("D" & NextCounter).Tag = Format(DateSerial(Year(d), Month(d), i), frmETRcalendar.ShortDateFormat)

            '--> Highlight the current day
            If i = thisDay And Month(d) = thisMonth And Year(d) = thisYear Then
                With Me.Controls("D" & NextCounter)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = CurDateColor
                    .ForeColor = CurDateForeColor
                End With
            Else '--> no highlight
                With Me.Controls("D" & NextCounter)
                    .BackStyle = fmBackStyleTransparent
                    .BackColor = MyBackColor
                    .ForeColor = MyForeColor
                End With
                
                '*** KEEP JUST IN CASE
'                Select Case Cal_theme
'                    Case CalendarThemes.ArcticBlue
'                        Me.Controls("D" & NextCounter).BackColor = CurDateColor
'                        Me.Controls("D" & NextCounter).ForeColor = RGB(0, 0, 0)
'                    Case Else
'                        Me.Controls("CB" & NextCounter).ForeColor = RGB(0, 0, 0)
'                End Select
                
                '********
            End If
    
            NextCounter = NextCounter + 1
        Next i
        
         '--> Populate days for the next month
        j = 1
        If NextCounter < 43 Then
            For i = NextCounter To 42
                With Me.Controls("D" & i)
                    .Caption = j
                    .Tag = Format(DateSerial(Year(dtNext), Month(dtNext), j), frmETRcalendar.ShortDateFormat)
                    .ForeColor = NotCurDateColor
                End With
                j = j + 1
            Next i
        End If
        
        'Populate days of previous month
        LastDay = Val(Format(dtLast, "dd"))
        If PrevCounter > 1 Then
            
            For i = PrevCounter To 1 Step -1
                With Me.Controls("D" & i)
                    .Caption = LastDay
                    .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmETRcalendar.ShortDateFormat)
                    .ForeColor = NotCurDateColor
                End With
                LastDay = LastDay - 1
            Next i
            
        ElseIf PrevCounter = 1 Then
        
            With Me.Controls("D1")
                .Caption = LastDay
                .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmETRcalendar.ShortDateFormat)
                .ForeColor = NotCurDateColor
            End With

        End If
        
        lblTitleCurMY.Caption = Format(d, "mmmm yyyy")
        
    End Sub
    
    '--> Show the months when user clicks on the date label
    Sub HiglightCurMonthControl()
         For i = 1 To 12
            
            If i = thisMonth Then
                With Me.Controls("M" & i)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = CurDateColor
                    .ForeColor = CurDateForeColor
                End With
            End If
         Next i
    End Sub
    
    '--> Show the details for the selected month
    Sub ShowSpecificMonth()
        lblTitleCurMY.Caption = Format(DateSerial(CurYear, CurMonth, 1), "mmm yyyy")
        MPmainDisplay.value = 0 'switch multipage back to 'Day' page
        PopulateCalendar DateSerial(CurYear, CurMonth, 1)
        Label5.Caption = 1
        lblUP.Visible = True
        lblDOWN.Visible = True
    End Sub
    
    '--> Handles the month to year multipage display
    Private Sub lblTitleCurMY_Click()
         Select Case Label5.Caption
            Case 1
           
                lblTitleCurMY.Caption = Split(lblTitleCurMY.Caption)(1)
                Label5.Caption = 2
                Me.MPmainDisplay.value = 1 '--> Switch active multipage
                HiglightCurMonthControl
                lblDOWN.Visible = False
                lblUP.Visible = False

            Case 2 '--> Prep & show year buttons
                
                lblDOWN.Visible = True
                lblUP.Visible = True
                Me.MPmainDisplay.value = 2 '--> Switch active multipage
                
                ToYr = Val(lblTitleCurMY.Caption)
                frmYr = ToYr - 11
                
                If frmYr < 1919 Then frmYr = 1919
                
                lblTitleCurMY.Caption = frmYr & " - " & ToYr
                Label5.Caption = 3
                
                For i = 1 To 12
                    Me.Controls("Y" & i).Caption = ""
                Next i
                
                For i = 12 To 1 Step -1
                    If Not ToYr < 1919 Then
                        With Me.Controls("Y" & i)
                            .Caption = ToYr
                            .Visible = True
                            
                            If ToYr = thisYear Then
                                With Me.Controls("Y" & i)
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = CurDateColor
                                    .ForeColor = CurDateForeColor
                                End With
                            End If
                            
                            ToYr = ToYr - 1
                        End With
                    End If
                Next i
                
                Label5.Caption = 3
            Case 3 'Do Nothing
         End Select
    End Sub
    
' Logicworkz 12/2019 ----------------------------------------------------------------
'--------- CALENDAR DAY LABEL "BUTTONS" BORDER MOUSE ENTRY/EXIT BEHAVIOR ------------
'------------------------------------------------------------------------------------
Sub NoBorder(SkipLabel As Byte, PreFix As String, ObjCnt As Byte)
    Dim d As Byte
    For d = 1 To ObjCnt
        If d <> SkipLabel Then
            With Me.Controls(PreFix & d)
                .BorderStyle = fmBorderStyleNone
            End With
        End If
    Next
End Sub

Private Sub D1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D1.BorderStyle = fmBorderStyleNone
    Else
        D1.BorderStyle = fmBorderStyleSingle
        NoBorder 1, "D", 42
    End If
End Sub
Private Sub D2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D2.BorderStyle = fmBorderStyleNone
    Else
        D2.BorderStyle = fmBorderStyleSingle
        NoBorder 2, "D", 42
    End If
End Sub
Private Sub D3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D3.BorderStyle = fmBorderStyleNone
    Else
        D3.BorderStyle = fmBorderStyleSingle
        NoBorder 3, "D", 42
    End If
End Sub
Private Sub D4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D4.BorderStyle = fmBorderStyleNone
    Else
        D4.BorderStyle = fmBorderStyleSingle
        NoBorder 4, "D", 42
    End If
End Sub
Private Sub D5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D5.BorderStyle = fmBorderStyleNone
    Else
        D5.BorderStyle = fmBorderStyleSingle
        NoBorder 5, "D", 42
    End If
End Sub
Private Sub D6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D6.BorderStyle = fmBorderStyleNone
    Else
        D6.BorderStyle = fmBorderStyleSingle
        NoBorder 6, "D", 42
    End If
End Sub
Private Sub D7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D7.BorderStyle = fmBorderStyleNone
    Else
        D7.BorderStyle = fmBorderStyleSingle
        NoBorder 7, "D", 42
    End If
End Sub
Private Sub D8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D8.BorderStyle = fmBorderStyleNone
    Else
        D8.BorderStyle = fmBorderStyleSingle
        NoBorder 8, "D", 42
    End If
End Sub
Private Sub D9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D9.BorderStyle = fmBorderStyleNone
    Else
        D9.BorderStyle = fmBorderStyleSingle
        NoBorder 9, "D", 42
    End If
End Sub
Private Sub D10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D10.BorderStyle = fmBorderStyleNone
    Else
        D10.BorderStyle = fmBorderStyleSingle
        NoBorder 10, "D", 42
    End If
End Sub
Private Sub D11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D11.BorderStyle = fmBorderStyleNone
    Else
        D11.BorderStyle = fmBorderStyleSingle
        NoBorder 11, "D", 42
    End If
End Sub
Private Sub D12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D12.BorderStyle = fmBorderStyleNone
    Else
        D12.BorderStyle = fmBorderStyleSingle
        NoBorder 12, "D", 42
    End If
End Sub
Private Sub D13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D13.BorderStyle = fmBorderStyleNone
    Else
        D13.BorderStyle = fmBorderStyleSingle
        NoBorder 13, "D", 42
    End If
End Sub
Private Sub D14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D14.BorderStyle = fmBorderStyleNone
    Else
        D14.BorderStyle = fmBorderStyleSingle
        NoBorder 14, "D", 42
    End If
End Sub
Private Sub D15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D15.BorderStyle = fmBorderStyleNone
    Else
        D15.BorderStyle = fmBorderStyleSingle
        NoBorder 15, "D", 42
    End If
End Sub
Private Sub D16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D16.BorderStyle = fmBorderStyleNone
    Else
        D16.BorderStyle = fmBorderStyleSingle
        NoBorder 16, "D", 42
    End If
End Sub
Private Sub D17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D17.BorderStyle = fmBorderStyleNone
    Else
        D17.BorderStyle = fmBorderStyleSingle
        NoBorder 17, "D", 42
    End If
End Sub
Private Sub D18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D18.BorderStyle = fmBorderStyleNone
    Else
        D18.BorderStyle = fmBorderStyleSingle
        NoBorder 18, "D", 42
    End If
End Sub
Private Sub D19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D19.BorderStyle = fmBorderStyleNone
    Else
        D19.BorderStyle = fmBorderStyleSingle
        NoBorder 19, "D", 42
    End If
End Sub
Private Sub D20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D20.BorderStyle = fmBorderStyleNone
    Else
        D20.BorderStyle = fmBorderStyleSingle
        NoBorder 20, "D", 42
    End If
End Sub
Private Sub D21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D21.BorderStyle = fmBorderStyleNone
    Else
        D21.BorderStyle = fmBorderStyleSingle
        NoBorder 21, "D", 42
    End If
End Sub
Private Sub D22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D22.BorderStyle = fmBorderStyleNone
    Else
        D22.BorderStyle = fmBorderStyleSingle
        NoBorder 22, "D", 42
    End If
End Sub
Private Sub D23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D23.BorderStyle = fmBorderStyleNone
    Else
        D23.BorderStyle = fmBorderStyleSingle
        NoBorder 23, "D", 42
    End If
End Sub
Private Sub D24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D24.BorderStyle = fmBorderStyleNone
    Else
        D24.BorderStyle = fmBorderStyleSingle
        NoBorder 24, "D", 42
    End If
End Sub
Private Sub D25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D25.BorderStyle = fmBorderStyleNone
    Else
        D25.BorderStyle = fmBorderStyleSingle
        NoBorder 25, "D", 42
    End If
End Sub
Private Sub D26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D26.BorderStyle = fmBorderStyleNone
    Else
        D26.BorderStyle = fmBorderStyleSingle
        NoBorder 26, "D", 42
    End If
End Sub
Private Sub D27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D27.BorderStyle = fmBorderStyleNone
    Else
        D27.BorderStyle = fmBorderStyleSingle
        NoBorder 27, "D", 42
    End If
End Sub
Private Sub D28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D28.BorderStyle = fmBorderStyleNone
    Else
        D28.BorderStyle = fmBorderStyleSingle
        NoBorder 28, "D", 42
    End If
End Sub
Private Sub D29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D29.BorderStyle = fmBorderStyleNone
    Else
        D29.BorderStyle = fmBorderStyleSingle
        NoBorder 29, "D", 42
    End If
End Sub
Private Sub D30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D30.BorderStyle = fmBorderStyleNone
    Else
        D30.BorderStyle = fmBorderStyleSingle
        NoBorder 30, "D", 42
    End If
End Sub
Private Sub D31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D31.BorderStyle = fmBorderStyleNone
    Else
        D31.BorderStyle = fmBorderStyleSingle
        NoBorder 31, "D", 42
    End If
End Sub
Private Sub D32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D32.BorderStyle = fmBorderStyleNone
    Else
        D32.BorderStyle = fmBorderStyleSingle
        NoBorder 32, "D", 42
    End If
End Sub
Private Sub D33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D33.BorderStyle = fmBorderStyleNone
    Else
        D33.BorderStyle = fmBorderStyleSingle
        NoBorder 33, "D", 42
    End If
End Sub
Private Sub D34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D34.BorderStyle = fmBorderStyleNone
    Else
        D34.BorderStyle = fmBorderStyleSingle
        NoBorder 34, "D", 42
    End If
End Sub
Private Sub D35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D35.BorderStyle = fmBorderStyleNone
    Else
        D35.BorderStyle = fmBorderStyleSingle
        NoBorder 35, "D", 42
    End If
End Sub
Private Sub D36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D36.BorderStyle = fmBorderStyleNone
    Else
        D36.BorderStyle = fmBorderStyleSingle
        NoBorder 36, "D", 42
    End If
End Sub
Private Sub D37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D37.BorderStyle = fmBorderStyleNone
    Else
        D37.BorderStyle = fmBorderStyleSingle
        NoBorder 37, "D", 42
    End If
End Sub
Private Sub D38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D38.BorderStyle = fmBorderStyleNone
    Else
        D38.BorderStyle = fmBorderStyleSingle
        NoBorder 38, "D", 42
    End If
End Sub
Private Sub D39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D39.BorderStyle = fmBorderStyleNone
    Else
        D39.BorderStyle = fmBorderStyleSingle
        NoBorder 39, "D", 42
    End If
End Sub
Private Sub D40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D40.BorderStyle = fmBorderStyleNone
    Else
        D40.BorderStyle = fmBorderStyleSingle
        NoBorder 40, "D", 42
    End If
End Sub
Private Sub D41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D41.BorderStyle = fmBorderStyleNone
    Else
        D41.BorderStyle = fmBorderStyleSingle
        NoBorder 41, "D", 42
    End If
End Sub
Private Sub D42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Day_Xmax Or X <= Day_Xmin Or Y <= Day_Ymin Or Y >= Day_Ymax Then
        D42.BorderStyle = fmBorderStyleNone
    Else
        D42.BorderStyle = fmBorderStyleSingle
        NoBorder 42, "D", 42
    End If
End Sub

' Logicworkz 12/2019 ----------------------------------------------------------------
'--------- CALENDAR MONTH LABEL "BUTTONS" BORDER MOUSE ENTRY/EXIT BEHAVIOR ----------
'------------------------------------------------------------------------------------
Private Sub M1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M1.BorderStyle = fmBorderStyleNone
    Else
        M1.BorderStyle = fmBorderStyleSingle
        NoBorder 1, "M", 12
    End If
End Sub
Private Sub M2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M2.BorderStyle = fmBorderStyleNone
    Else
        M2.BorderStyle = fmBorderStyleSingle
        NoBorder 2, "M", 12
    End If
End Sub
Private Sub M3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M3.BorderStyle = fmBorderStyleNone
    Else
        M3.BorderStyle = fmBorderStyleSingle
        NoBorder 3, "M", 12
    End If
End Sub
Private Sub M4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M4.BorderStyle = fmBorderStyleNone
    Else
        M4.BorderStyle = fmBorderStyleSingle
        NoBorder 4, "M", 12
    End If
End Sub
Private Sub M5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M5.BorderStyle = fmBorderStyleNone
    Else
        M5.BorderStyle = fmBorderStyleSingle
        NoBorder 5, "M", 12
    End If
End Sub
Private Sub M6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M6.BorderStyle = fmBorderStyleNone
    Else
        M6.BorderStyle = fmBorderStyleSingle
        NoBorder 6, "M", 12
    End If
End Sub
Private Sub M7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M7.BorderStyle = fmBorderStyleNone
    Else
        M7.BorderStyle = fmBorderStyleSingle
        NoBorder 7, "M", 12
    End If
End Sub
Private Sub M8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M8.BorderStyle = fmBorderStyleNone
    Else
        M8.BorderStyle = fmBorderStyleSingle
        NoBorder 8, "M", 12
    End If
End Sub
Private Sub M9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M9.BorderStyle = fmBorderStyleNone
    Else
        M9.BorderStyle = fmBorderStyleSingle
        NoBorder 9, "M", 12
    End If
End Sub
Private Sub M10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M10.BorderStyle = fmBorderStyleNone
    Else
        M10.BorderStyle = fmBorderStyleSingle
        NoBorder 10, "M", 12
    End If
End Sub
Private Sub M11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M11.BorderStyle = fmBorderStyleNone
    Else
        M11.BorderStyle = fmBorderStyleSingle
        NoBorder 11, "M", 12
    End If
End Sub
Private Sub M12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        M12.BorderStyle = fmBorderStyleNone
    Else
        M12.BorderStyle = fmBorderStyleSingle
        NoBorder 12, "M", 12
    End If
End Sub
' Logicworkz 12/2019 ----------------------------------------------------------------
'--------- MISC LABEL "BUTTON" BORDER MOUSE ENTRY/EXIT BEHAVIOR ----------
'------------------------------------------------------------------------------------
Private Sub lblTitleCurMY_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= 68 Or X <= 2 Or Y <= 2 Or Y >= 10 Then
        lblTitleCurMY.ForeColor = MyForeColor
    Else
        lblTitleCurMY.ForeColor = RGB(73, 255, 60)
    End If
End Sub

Private Sub lblUP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= 20 Or X <= 2 Or Y <= 2 Or Y >= 10 Then
        lblUP.ForeColor = MyForeColor
    Else
        lblUP.ForeColor = RGB(73, 255, 60)
    End If
End Sub
Private Sub lblDOWN_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= 20 Or X <= 2 Or Y <= 2 Or Y >= 10 Then
        lblDOWN.ForeColor = MyForeColor
    Else
        lblDOWN.ForeColor = RGB(73, 255, 60)
    End If
End Sub
Private Sub lblUnload_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     If X >= 20 Or X <= 2 Or Y <= 2 Or Y >= 10 Then
        lblUnload.ForeColor = MyForeColor
    Else
        lblUnload.ForeColor = RGB(73, 255, 60)
    End If
End Sub
Private Sub lblThemes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     If X >= 20 Or X <= 2 Or Y <= 2 Or Y >= 10 Then
        lblThemes.ForeColor = MyForeColor
    Else
        lblThemes.ForeColor = RGB(73, 255, 60)
    End If
End Sub
' Logicworkz 12/2019 ----------------------------------------------------------------
'--------- CALENDAR MONTH HEADINGS MOUSE ENTRY/EXIT BEHAVIOR ----------
'------------------------------------------------------------------------------------
Private Sub WD1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
Private Sub WD7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim d As Byte
    For d = 1 To 42
        With Me.Controls("D" & d)
            .BorderStyle = fmBorderStyleNone
        End With
    Next
End Sub
' Logicworkz 12/2019 ----------------------------------------------------------------
'--------- CALENDAR YR LABEL "BUTTON" BORDER MOUSE ENTRY/EXIT BEHAVIOR ----------
'------------------------------------------------------------------------------------
Private Sub Y1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y1.BorderStyle = fmBorderStyleNone
    Else
        Y1.BorderStyle = fmBorderStyleSingle
        NoBorder 1, "Y", 12
    End If
End Sub
Private Sub Y2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y2.BorderStyle = fmBorderStyleNone
    Else
        Y2.BorderStyle = fmBorderStyleSingle
        NoBorder 2, "Y", 12
    End If
End Sub
Private Sub Y3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y3.BorderStyle = fmBorderStyleNone
    Else
        Y3.BorderStyle = fmBorderStyleSingle
        NoBorder 3, "Y", 12
    End If
End Sub
Private Sub Y4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y4.BorderStyle = fmBorderStyleNone
    Else
        Y4.BorderStyle = fmBorderStyleSingle
        NoBorder 4, "Y", 12
    End If
End Sub
Private Sub Y5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y5.BorderStyle = fmBorderStyleNone
    Else
        Y5.BorderStyle = fmBorderStyleSingle
        NoBorder 5, "Y", 12
    End If
End Sub
Private Sub Y6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y6.BorderStyle = fmBorderStyleNone
    Else
        Y6.BorderStyle = fmBorderStyleSingle
        NoBorder 6, "Y", 12
    End If
End Sub
Private Sub Y7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y7.BorderStyle = fmBorderStyleNone
    Else
        Y7.BorderStyle = fmBorderStyleSingle
        NoBorder 7, "Y", 12
    End If
End Sub
Private Sub Y8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y8.BorderStyle = fmBorderStyleNone
    Else
        Y8.BorderStyle = fmBorderStyleSingle
        NoBorder 8, "Y", 12
    End If
End Sub
Private Sub Y9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y9.BorderStyle = fmBorderStyleNone
    Else
        Y9.BorderStyle = fmBorderStyleSingle
        NoBorder 9, "Y", 12
    End If
End Sub
Private Sub Y10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y10.BorderStyle = fmBorderStyleNone
    Else
        Y10.BorderStyle = fmBorderStyleSingle
        NoBorder 10, "Y", 12
    End If
End Sub
Private Sub Y11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y11.BorderStyle = fmBorderStyleNone
    Else
        Y11.BorderStyle = fmBorderStyleSingle
        NoBorder 11, "Y", 12
    End If
End Sub
Private Sub Y12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If X >= Mo_Xmax Or X <= Mo_Xmin Or Y <= Mo_Ymin Or Y >= Mo_Ymax Then
        Y12.BorderStyle = fmBorderStyleNone
    Else
        Y12.BorderStyle = fmBorderStyleSingle
        NoBorder 12, "Y", 12
    End If
End Sub


'this slows the border appearance down but prevent a fast mouse from leaving a border visible after exit
'Private Sub FrameMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'
'    Dim lblCtrl As control
'    i = 0
'    For Each lblCtrl In Me.Controls
'        If TypeOf lblCtrl Is MSForms.Label Then
'            lblCtrl.BorderStyle = fmBorderStyleNone
'        End If
'    Next
'End Sub

