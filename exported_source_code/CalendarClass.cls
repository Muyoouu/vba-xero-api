VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Credits: Siddharth Rout, logicworkz
' Source: https://stackoverflow.com/questions/54650417/how-can-i-create-a-calendar-input-in-vba-excel
    '
    '~~> This section is used for handling Commandbutton Control Array
    '
    Public WithEvents CommandButtonEvents As MSForms.Label '****** JCR
Attribute CommandButtonEvents.VB_VarHelpID = -1
    
    '~~> Unload the form when the user presses Escape
'    Private Sub CommandButtonEvents_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'        If Not f Is Nothing Then If KeyAscii = 27 Then Unload f
'    End Sub
    
    '~~> This section delas with showing/displaying controls
    '~~> and updating different labels
    Private Sub CommandButtonEvents_Click()
    
        '--> rec main subroutine entry point
        If CommandButtonEvents.Tag <> "" Then
            
            WebHelpers.LogDebug "User selected date: " & CommandButtonEvents.Tag, "CalendarClass"
            With frmETRcalendar
                .UserSelectedDateStr = CommandButtonEvents.Tag
                .hide
            End With
        End If
        
        If Left(CommandButtonEvents.name, 1) = "Y" Then
            If Len(Trim(CommandButtonEvents.Caption)) <> 0 Then
                CurYear = Val(CommandButtonEvents.Caption)

                With frmETRcalendar
                    .MPmainDisplay.value = 1 'yr selected, go back to month display
                
                    .lblTitleCurMY.Caption = CurYear
                    .Label5.Caption = 2
                    
                    .lblUP.Visible = False
                    .lblDOWN.Visible = False

                End With
            End If
            
        ElseIf Left(CommandButtonEvents.name, 1) = "M" Then
            
            Select Case CommandButtonEvents.Caption
                Case "JAN": CurMonth = 1
                Case "FEB": CurMonth = 2
                Case "MAR": CurMonth = 3
                Case "APR": CurMonth = 4
                Case "MAY": CurMonth = 5
                Case "JUN": CurMonth = 6
                Case "JUL": CurMonth = 7
                Case "AUG": CurMonth = 8
                Case "SEP": CurMonth = 9
                Case "OCT": CurMonth = 10
                Case "NOV": CurMonth = 11
                Case "DEC": CurMonth = 12
            End Select

            frmETRcalendar.ShowSpecificMonth
        End If
    End Sub
