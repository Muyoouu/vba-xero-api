Attribute VB_Name = "CalendarModule"
' Credits: Siddharth Rout, logicworkz
' Source: https://stackoverflow.com/questions/54650417/how-can-i-create-a-calendar-input-in-vba-excel
    
    Option Explicit
    
    Public Const GWL_STYLE = -16
    Public Const WS_CAPTION = &HC00000
       
    #If VBA7 Then
        #If Win64 Then
            Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
            
            Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
        #Else
            Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
            
            Private Declare Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
        #End If
        
        Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
        
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
        
        Private Declare PtrSafe Function SetTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As LongPtr
    
        Public Declare PtrSafe Function KillTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr
        
        Public TimerID As LongPtr
        
        Dim lngWindow As LongPtr, lFrmHdl As LongPtr
    #Else
    
        Public Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long) As Long
        
        Public Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
        Public Declare Function DrawMenuBar _
        Lib "user32" (ByVal hwnd As Long) As Long
        
        Public Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    
        Public Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
        
        Public Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
        
        Public TimerID As Long
        Dim lngWindow As Long, lFrmHdl As Long
    #End If
    
    Public TimerSeconds As Single, tim As Boolean
    Public CurMonth As Integer, CurYear As Integer
    Public frmYr As Integer, ToYr As Integer
    
    ' DEBUG MUY: not used and it caused conflicts with other modules
    'Public F As frmETRcalendar
    
    Enum CalendarThemes
        Venom = 0
        MartianRed = 1
        ArcticBlue = 2
        Greyscale = 3
    End Enum
    
    Public Const Day_Xmax = 20
    Public Const Day_Xmin = 2
    Public Const Day_Ymax = 15.75
    Public Const Day_Ymin = 2
    
    Public Const Mo_Xmax = 38
    Public Const Mo_Xmin = 2
    Public Const Mo_Ymax = 15.75
    Public Const Mo_Ymin = 2
    
    Public MyBackColor As Long, MyForeColor As Long, CurDateColor As Long, CurDateForeColor As Long, NotCurDateColor As Long
    
    Sub Launch() '(control As IRibbonControl)
        
        With frmETRcalendar
            .Caltheme = CalendarThemes.Venom
            .LongDateFormat = "dddd, mmmm dd" '"mmmm dddd dd, yyyy" '"dddd dd. mmmm yyyy" ' etc
            .ShortDateFormat = "mm/dd/yyyy" 'or "d/m/y" etc
            .show
        End With
        
    End Sub
    
    '~~> Hide the title bar of the userform
    Sub HideTitleBar(frm As Object)
        #If VBA7 Then
            Dim lngWindow As LongPtr, lFrmHdl As LongPtr
            lFrmHdl = FindWindow(vbNullString, frm.Caption)
            lngWindow = GetWindowLongPtr(lFrmHdl, GWL_STYLE)
            lngWindow = lngWindow And (Not WS_CAPTION)
            Call SetWindowLongPtr(lFrmHdl, GWL_STYLE, lngWindow)
            Call DrawMenuBar(lFrmHdl)
        #Else
            Dim lngWindow As Long, lFrmHdl As Long
            lFrmHdl = FindWindow(vbNullString, frm.Caption)
            lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
            lngWindow = lngWindow And (Not WS_CAPTION)
            Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
            Call DrawMenuBar(lFrmHdl)
        #End If
    End Sub

    Sub StartTimer()
        '~~ Set the timer for 1 second
        TimerSeconds = 1
        TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
    End Sub

    Sub EndTimer()
        On Error Resume Next
        KillTimer 0&, TimerID
    End Sub
        
    '~~> Update Time
    #If VBA7 And Win64 Then ' 64 bit Excel under 64-bit windows  ' Use LongLong and LongPtr
        Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As LongLong, _
        ByVal nIDEvent As LongPtr, ByVal dwTimer As LongLong)
            frmETRcalendar.lblTitleClock.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
            frmETRcalendar.lblTitleAMPM.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
        End Sub
    #ElseIf VBA7 Then ' 64 bit Excel in all environments
        Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
        ByVal nIDEvent As LongPtr, ByVal dwTimer As Long)
            frmETRcalendar.lblTitleClock.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
            frmETRcalendar.lblTitleAMPM.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
        End Sub
    #End If
