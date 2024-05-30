Attribute VB_Name = "XeroAPICall"
' XeroAPICall v1.0.0
' @author musayohanes00@gmail.com
' https://github.com/Muyoouu/vba-xero-api
'
' Xero accounting API calls
' Docs: https://developer.xero.com/documentation/api/accounting/overview

Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' Provide Client ID and Secret through constant variable or user will be prompted
Private Const cXEROCLIENTID As String = ""
Private Const cXEROCLIENTSECRET As String = ""
' Prefix name to the output sheet for the report
Private Const ReportOutputSheet As String = "P&L_Report_"

' Used for naming JSON file output if any
Private pLastOutputSheetName As String
' Cache WebClient for API calls
Private pXeroClient As WebClient
Private pXeroClientId As String
Private pXeroClientSecret As String

' --------------------------------------------- '
' Private Properties and Methods
' --------------------------------------------- '

' Property to get API Client ID
Private Property Get XeroClientId() As String
    If pXeroClientId = "" Then
        If cXEROCLIENTID <> "" Then
            pXeroClientId = cXEROCLIENTID
        Else
            Dim InpBxResponse As String
            InpBxResponse = InputBox("Please Enter Xero API Client ID", "Xero Report Generator - Microsoft Excel")
            If InpBxResponse <> "" Then
                pXeroClientId = InpBxResponse
            Else
                Err.Raise 11041 + vbObjectError, "XeroAPICall.ClientIdInputBox", "User did not provide Xero API Client ID"
            End If
        End If
    End If
    
    XeroClientId = pXeroClientId
End Property

' Property to get API Client Secret
Private Property Get XeroClientSecret() As String
    If pXeroClientSecret = "" Then
        If cXEROCLIENTSECRET <> "" Then
            pXeroClientSecret = cXEROCLIENTSECRET
        Else
            Dim InpBxResponse As String
            InpBxResponse = InputBox("Please Enter Xero API Client Secret", "Xero Report Generator - Microsoft Excel")
            If InpBxResponse <> "" Then
                pXeroClientSecret = InpBxResponse
            Else
                Err.Raise 11041 + vbObjectError, "XeroAPICall.ClientSecretInputBox", "User did not provide Xero API Client Secret"
            End If
        End If
    End If
    
    XeroClientSecret = pXeroClientSecret
End Property

' Setup client and authenticator (cached between requests)
Private Property Get XeroClient() As WebClient
    If pXeroClient Is Nothing Then
        ' Create client with base url that is appended to all requests
        Set pXeroClient = New WebClient
        pXeroClient.BaseUrl = "https://api.xero.com/"
        
        ' Use the custom made XeroAuthenticator
        ' - Automatically uses Xero's OAuth2 approach including login screen
        Dim Auth As XeroAuthenticator
        Set Auth = New XeroAuthenticator
        Auth.Setup CStr(XeroClientId), CStr(XeroClientSecret)
        ' Make sure to request refresh token with 'offline_access' scope included
        Auth.AddScope "offline_access"
        Auth.AddScope "accounting.reports.read"
        
        Set pXeroClient.Authenticator = Auth
    End If
    
    Set XeroClient = pXeroClient
End Property

' Property to set XeroClient
Private Property Set XeroClient(Client As WebClient)
    Set pXeroClient = Client
End Property

' Load SelectReportForm and pass user selections
Private Function SelectReport(Request As WebRequest) As WebRequest
    On Error GoTo ApiCall_Cleanup

    ' Initialize form
    Dim SelectForm1 As SelectReportForm
    Set SelectForm1 = New SelectReportForm
    
    ' Show form to user
    SelectForm1.show
    
    ' Check if user canceled the form
    If SelectForm1.UserCancel Then
        ' Notify user and raise error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly
        Err.Raise 11040 + vbObjectError, "SelectReportForm", "User canceled selection form"
    End If
    
    ' Change dates format following API docs and assign to request params
    Dim fromDate As Date
    Dim toDate As Date
    fromDate = DateSerial(CInt(Right(SelectForm1.TextBox1.value, 4)), CInt(Left(SelectForm1.TextBox1.value, 2)), CInt(Mid(SelectForm1.TextBox1.value, 4, 2)))
    toDate = DateSerial(CInt(Right(SelectForm1.TextBox2.value, 4)), CInt(Left(SelectForm1.TextBox2.value, 2)), CInt(Mid(SelectForm1.TextBox2.value, 4, 2)))
    
    Request.AddQuerystringParam "fromDate", Format(fromDate, "yyyy-mm-dd")
    Request.AddQuerystringParam "toDate", Format(toDate, "yyyy-mm-dd")
    
    ' Success and return
    Set SelectReport = Request

ApiCall_Cleanup:
    ' Unload when everything is finished
    If Not SelectForm1 Is Nothing Then
        Unload SelectForm1
    End If
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the report selection process." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.SelectReport", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAPICall.SelectReport", auth_ErrorDescription
    End If
End Function

' Send GET request to API for P&L report
Private Function GetPnLReport() As Dictionary
    On Error GoTo ApiCall_Cleanup

    ' Initialize form
    Dim ReportRequest As WebRequest
    Set ReportRequest = New WebRequest
    
    ' Prepare report request
    ReportRequest.Resource = "api.xro/2.0/Reports/ProfitAndLoss"
    ReportRequest.Method = WebMethod.HttpGet
    ReportRequest.RequestFormat = WebFormat.FormUrlEncoded
    ReportRequest.ResponseFormat = WebFormat.Json
    
    ' Let user select report details and period
    Set ReportRequest = SelectReport(ReportRequest)
    
    ' Sent get request and receive response
    Dim ReportResponse As WebResponse
    Set ReportResponse = XeroClient.Execute(ReportRequest)
    
    ' Success and return
    If ReportResponse.StatusCode = WebStatusCode.Ok Then
        Set GetPnLReport = ReportResponse.Data
    Else
        Err.Raise 11041 + vbObjectError, "XeroAPICall.GetPnLReport", _
            ReportResponse.StatusCode & ": " & ReportResponse.Content
    End If
    
ApiCall_Cleanup:
    ' Terminate when everything is finished
    Set ReportRequest = Nothing
    Set ReportResponse = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving a profit and loss report." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.GetPnLReport", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAPICall.GetPnLReport", auth_ErrorDescription
    End If
End Function

' Parse response data (JSON in form of Dictionary) and load it to sheet
Private Sub LoadReportToSheet(GetReportData As Dictionary, Optional SheetName As String = ReportOutputSheet)
    On Error GoTo ApiCall_Cleanup

    ' GetReportData should contain JSON object obtained from API call
    Dim report  As Dictionary
    Set report = GetReportData("Reports")(1)
    
    Dim reportTitles As Collection
    Set reportTitles = report("ReportTitles")
    
    Dim rows As Collection
    Set rows = report("Rows")
    
    ' To track Excel sheet rows index
    Dim rowIndex As Long
    rowIndex = 1 ' Start at first row

    ' Sheet to load the JSON into
    Dim sh As Worksheet
    Dim sheetIndex As Integer
    sheetIndex = 0
    
    ' Create new sheets
    Set sh = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    
    ' Write headers
    With sh
        ' Loop through ReportTitles to add each line to the sheet
        Dim reportTitle As Variant
        For Each reportTitle In reportTitles
            ' Fill titles into sheet
            If rowIndex = 3 Then
                ' Add words for the third row title - the report period
                .Cells(rowIndex, 1).value = "For the period of " & reportTitle
            Else
                .Cells(rowIndex, 1).value = reportTitle
            End If
            ' Format the font styles
            With .Cells(rowIndex, 1).Font
                If rowIndex = 1 Then
                    .Bold = True
                    .Size = 14
                Else
                    .Size = 12
                End If
            End With
            rowIndex = rowIndex + 1
        Next reportTitle
        ' Add a blank row after titles
        rowIndex = rowIndex + 1

        ' Add the Account and Date headers
        .Cells(rowIndex, 1).value = "Account"
        
        ' Re-format the report dates
        Dim dates() As String
        Dim dateFormat As String
        dates = Split(reportTitles(3), " to ")
        ' Conditionally format the date header
        dateFormat = "d mmm"
        If Year(CDate(dates(0))) <> Year(CDate(dates(1))) Then
            dateFormat = dateFormat & " yyyy"
        End If
        ' Assign the formatted date
        .Cells(rowIndex, 2).value = Format(CDate(dates(0)), dateFormat) & "-" & Format(CDate(dates(1)), "d mmm yyyy")
        ' Add end date of the report period for sheet name
        SheetName = SheetName & UCase(Format(CDate(dates(1)), "dmmmyy"))
        
        ' Format the cells styles
        With .Range(.Cells(rowIndex, 1), .Cells(rowIndex, 2))
            .Font.Bold = True
            .Font.Size = 10
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        ' Add a blank row after account and date headers
        rowIndex = rowIndex + 2

        Dim section As Dictionary
        Dim innerRows As Collection
        Dim row As Dictionary
        
        For Each section In rows
            If section("RowType") = "Section" Then
                If section("Title") <> "" Then
                    .Cells(rowIndex, 1).value = section("Title")
                    ' Styling for section header
                    With .Range(.Cells(rowIndex, 1), .Cells(rowIndex, 2))
                        .Font.Bold = True
                        .Font.Size = 10
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).ColorIndex = 1
                    End With
                    rowIndex = rowIndex + 1
                End If

                Set innerRows = section("Rows")
                For Each row In innerRows
                    .Cells(rowIndex, 1).value = row("Cells")(1)("Value")
                    .Cells(rowIndex, 2).value = row("Cells")(2)("Value")
                    ' Format number values
                    .Cells(rowIndex, 2).NumberFormat = "#,##0.00;(#,##0.00)"
                
                    ' Styling for each row
                    With .Range(.Cells(rowIndex, 1), .Cells(rowIndex, 2))
                        .Font.Size = 9
                        If row("RowType") = "SummaryRow" Then
                            .Font.Bold = True
                        ElseIf section("Title") = "" Then
                            .Font.Bold = True
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Borders(xlEdgeBottom).ColorIndex = 1
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Borders(xlEdgeTop).ColorIndex = 1
                            .Interior.Color = RGB(242, 242, 242)
                        End If
                    End With
                    
                    rowIndex = rowIndex + 1
                Next row

                ' Add a blank row after each section
                rowIndex = rowIndex + 1
            End If
        Next section
        ' Set font
        .Range(.Cells(1, 1), .Cells(rowIndex, 2)).Font.name = "Arial"
        ' Autofit for better layout
        .Range(.Cells(5, 1), .Cells(rowIndex, 2)).Columns.AutoFit
    End With
    ' Naming sheet
    ' Check available name
    If WebHelpers.WorksheetExists(SheetName, ThisWorkbook) Then
        Do While WebHelpers.WorksheetExists(SheetName & "_" & sheetIndex, ThisWorkbook)
            sheetIndex = sheetIndex + 1
        Loop
        SheetName = SheetName & "_" & sheetIndex
    End If
    sh.name = SheetName
    ' Turn off view gridlines
    WebHelpers.TurnOffGridLines sh
    
    ' Easily save the name of the sheet
    pLastOutputSheetName = SheetName
    
ApiCall_Cleanup:
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while loading report to sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.LoadReportToSheet", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAPICall.LoadReportToSheet", auth_ErrorDescription
    End If
End Sub

' --------------------------------------------- '
' Execution
' --------------------------------------------- '

' Call Login procedures (for user interface button)
Public Sub Login_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve pre-set authenticator object
    Dim Auth As XeroAuthenticator
    Set Auth = XeroClient.Authenticator
    Set XeroClient.Authenticator = Nothing
    
    ' Logout and clears cache for current session
    Auth.Logout
    
    ' Login
    Auth.Login
    
    ' Return auth reference to XeroClient
    Set XeroClient.Authenticator = Auth
    Set Auth = Nothing
    
ApiCall_Cleanup:
    ' Rethrow error
    If Err.Number <> 0 Then
        ' Clean up if error happened
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Error handling
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the login process." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.Login_Click", 11041 + vbObjectError
        ' Notify user
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub

' Call report generation procedures (for user interface button)
Public Sub GenerateReport_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve report from API
    Dim ReportDict As Dictionary
    Set ReportDict = GetPnLReport
        
    ' Parse and load the report into a sheet
    LoadReportToSheet ReportDict
    MsgBox "Report successfully generated on sheet: " & vbNewLine & pLastOutputSheetName, vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ApiCall_Cleanup:
    ' Rethrow error
    If Err.Number <> 0 Then
        ' Clean up if error happened
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Error handling
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while generating report." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.GenerateReport_Click", 11041 + vbObjectError
        ' Notify user
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub

' Clear all saved tokens and Xero organizations/tenants ID (for user interface button)
Public Sub ClearCache_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Confirm user action
    Dim msgBoxResponse As VbMsgBoxResult
    msgBoxResponse = MsgBox("This action will clear saved tokens (access) and Xero organization IDs. You will be required to log in for the next request to generate a report." & _
        vbNewLine & vbNewLine & "Proceed to clears cache?", vbExclamation + vbYesNo, "Xero Report Generator - Microsoft Excel")
    
    Select Case msgBoxResponse
        Case vbYes
            ' Retrieve pre-set authenticator object
            Dim Auth As XeroAuthenticator
            Set Auth = XeroClient.Authenticator
            Set XeroClient.Authenticator = Nothing
            
            ' Clears all cache
            Auth.ClearAllCache isClearTenant:=True, isClearToken:=True
            
            ' Clears current session tokens cache
            Auth.Logout
            
            ' Return auth reference to XeroClient
            Set XeroClient.Authenticator = Auth
            Set Auth = Nothing
            
        Case vbNo
            Exit Sub
    End Select

ApiCall_Cleanup:
    ' Rethrow error
    If Err.Number <> 0 Then
        ' Clean up if error happened
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Error handling
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.ClearCache_Click", 11041 + vbObjectError
        ' Notify user
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub
