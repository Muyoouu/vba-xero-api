VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectXeroOrgForm 
   Caption         =   "Select Xero Organization"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "SelectXeroOrgForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectXeroOrgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pOrgList As Variant
Public UserRequestUpdate As Boolean
Public UserCancel As Boolean

Public Property Let OrgList(XeroTenantsNames As Variant)
    pOrgList = XeroTenantsNames
End Property

Private Sub cmdbCancel_Click()
    ' Change value, for further process outside form
    UserCancel = True
    Me.hide
End Sub

Private Sub cmdbSubmit_Click()
    If ComboBox1.value = "" Then
        ' User can't submit blank value
        MsgBox "You can't submit a blank value, pick one from the list!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ComboBox1.SetFocus
    ElseIf ComboBox1.ListIndex = -1 Then
        ' User can't submit unknown values, must match the list
        MsgBox "You can't submit names other than from the list, pick one from the list!", vbExclamation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ComboBox1.SetFocus
    Else
        ' Submit success
        MsgBox ComboBox1.value & " organization is selected." & vbNewLine & "Processing into report generation.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Me.hide
    End If
End Sub

Private Sub cmdbUpdate_Click()
    ' Change value, for further process outside form
    UserRequestUpdate = True
    ' Switch controls visibility and un-enabled
    cmdbUpdate.Enabled = False
    Label2.Caption = "Still cannot find your Xero Organization?" & vbNewLine & _
                     "You might need to authorize the connection first. Cancel and try to re-login!"
    ' Hide the form from user
    Me.hide
End Sub

Private Sub UserForm_Activate()
    If Not IsEmpty(pOrgList) Then
        ComboBox1.List = pOrgList
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Label set up
    Label1.Visible = True
    Label2.Visible = True
    Label2.Caption = "If you cannot find your Xero organization, try updating."

    ' ComboBox set up
    ComboBox1.Enabled = True
    ComboBox1.Clear
    
    ' Command button set up
    cmdbUpdate.Visible = True
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
