VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectReportForm 
   Caption         =   "Select Report Details"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4410
   OleObjectBlob   =   "SelectReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserCancel As Boolean

Private Sub LblClndr1_Click()
    CalendarModule.Launch
    
    If Not frmETRcalendar Is Nothing Then
        If frmETRcalendar.UserSelectedDateStr <> "" Then
            TextBox1.value = frmETRcalendar.UserSelectedDateStr
            TextBox1.BackColor = RGB(255, 255, 255)
        End If
        Unload frmETRcalendar
    End If
End Sub

Private Sub LblClndr2_Click()
    CalendarModule.Launch
    
    If Not frmETRcalendar Is Nothing Then
        If frmETRcalendar.UserSelectedDateStr <> "" Then
            TextBox2.value = frmETRcalendar.UserSelectedDateStr
            TextBox2.BackColor = RGB(255, 255, 255)
        End If
        Unload frmETRcalendar
    End If
End Sub

Private Sub UserForm_Initialize()
    ' TextBox set up
    TextBox1.value = ""
    TextBox1.Enabled = False
    TextBox1.BackColor = RGB(255, 255, 255)
    TextBox2.value = ""
    TextBox2.Enabled = False
    TextBox2.BackColor = RGB(255, 255, 255)
    
    ' ComboBox set up
    ComboBox1.Enabled = False
    ComboBox1.value = "Profit & Loss Report"
    
    ' Command button set up
    cmdbCancel.Visible = True
    cmdbSubmit.Visible = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Me.hide
        ' Change value, for further process outside form
        UserCancel = True
    End If
End Sub

Private Sub cmdbCancel_Click()
    ' Change value, for further process outside form
    UserCancel = True
    Me.hide
End Sub

Private Sub cmdbSubmit_Click()
    Dim InvalidSubmit As Boolean
    InvalidSubmit = False
    
    Dim fromDate As Date
    Dim toDate As Date
    
    If ComboBox1.value = "" Then
        ' User can't submit blank value
        MsgBox "Please select a report type from the list!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ComboBox1.SetFocus
        InvalidSubmit = True
    End If
    
    If TextBox1.value = "" Then
        ' User can't submit blank value
        MsgBox "Please select the start date of the reporting period using the calendar icon!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        TextBox1.BackColor = RGB(255, 230, 230)
        InvalidSubmit = True
    Else
        fromDate = DateSerial(CInt(Right(TextBox1.value, 4)), CInt(Left(TextBox1.value, 2)), CInt(Mid(TextBox1.value, 4, 2)))
        If fromDate > Date Then
            ' User can't submit a future date for reporting period
            MsgBox "The start date you selected for the reporting period is invalid. It must not be a future date.", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            TextBox1.BackColor = RGB(255, 230, 230)
            InvalidSubmit = True
        End If
    End If
    
    If TextBox2.value = "" Then
        ' User can't submit blank value
        MsgBox "Please select the end date of the reporting period using the calendar icon!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        TextBox2.BackColor = RGB(255, 230, 230)
        InvalidSubmit = True
    Else
        toDate = DateSerial(CInt(Right(TextBox2.value, 4)), CInt(Left(TextBox2.value, 2)), CInt(Mid(TextBox2.value, 4, 2)))
        If toDate > Date Then
            ' User can't submit a future date for reporting period
            MsgBox "The end date you selected for the reporting period is invalid. It must not be a future date.", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            TextBox2.BackColor = RGB(255, 230, 230)
            InvalidSubmit = True
        End If
    End If
    
    If Not InvalidSubmit Then
        If Not fromDate < toDate Then
            ' User can't submit blank value
            MsgBox "The reporting period you selected is invalid. The start date must be earlier than the end date.", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            TextBox1.BackColor = RGB(255, 230, 230)
            TextBox2.BackColor = RGB(255, 230, 230)
        Else
            ' Submit success
            MsgBox ComboBox1.value & " will be generated with period starting from " & _
                TextBox1.value & " to " & TextBox2.value, vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            Me.hide
        End If
    End If
End Sub

