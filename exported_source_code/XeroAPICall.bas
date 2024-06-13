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

' Provide the Xero client ID and client secret through these constants.
' Leave these constants empty to be prompted for the values during runtime.
Private Const cXEROCLIENTID As String = ""
Private Const cXEROCLIENTSECRET As String = ""

' Prefix used for naming the output sheet where the Profit and Loss report will be generated.
Private Const ReportOutputSheet As String = "P&L_Report_"

' Used for naming the JSON file output, if any.
Private pLastOutputSheetName As String

' WebClient instance used for making API calls to Xero.
Private pXeroClient As WebClient

' Xero client ID and client secret values used for authentication.
Private pXeroClientId As String
Private pXeroClientSecret As String

' --------------------------------------------- '
' Private Properties and Methods
' --------------------------------------------- '

''
' Retrieves the Xero API client ID.
' If the client ID is not provided through the 'cXEROCLIENTID' constant, the user is prompted to enter the client ID.
'
' @property XeroClientId
' @type {String}
' @return {String} The Xero API client ID.
''
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

''
' Retrieves the Xero API client secret.
' If the client secret is not provided through the 'cXEROCLIENTSECRET' constant, the user is prompted to enter the client secret.
'
' @property XeroClientSecret
' @type {String}
' @return {String} The Xero API client secret.
''
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

''
' Initializes and returns a WebClient instance configured for making API calls to Xero.
'
' @property XeroClient
' @type {WebClient}
' @return {WebClient} The configured WebClient instance.
'
' The WebClient instance is set up with the following configurations:
' - Base URL set to 'https://api.xero.com/'
' - Authenticator set to an instance of the 'XeroAuthenticator' class, which handles Xero's OAuth2 authentication flow.
' - The 'offline_access' and 'accounting.reports.read' scopes are requested during the authentication process.
'
' The WebClient instance is cached and reused between requests.
''
Private Property Get XeroClient() As WebClient
    If pXeroClient Is Nothing Then
        ' Create a new WebClient instance with the base URL
        Set pXeroClient = New WebClient
        pXeroClient.BaseUrl = "https://api.xero.com/"
        
        ' Set up the 'XeroAuthenticator' instance for OAuth2 authentication
        Dim Auth As XeroAuthenticator
        Set Auth = New XeroAuthenticator
        Auth.Setup CStr(XeroClientId), CStr(XeroClientSecret)
        
        ' Request the 'offline_access' and 'accounting.reports.read' scopes
        Auth.AddScope "offline_access"
        Auth.AddScope "accounting.reports.read"
        
        ' Set the 'XeroAuthenticator' instance as the authenticator for the WebClient
        Set pXeroClient.Authenticator = Auth
    End If
    
    Set XeroClient = pXeroClient
End Property

''
' Sets the WebClient instance used for making API calls to Xero.
'
' @property XeroClient
' @type {WebClient}
' @param {WebClient} Client - The WebClient instance to set.
''
Private Property Set XeroClient(Client As WebClient)
    Set pXeroClient = Client
End Property

''
' Displays a user form that allows the user to select the report parameters (date range) for the Xero API request.
'
' @method SelectReport
' @param {WebRequest} Request - The WebRequest object to which the selected report parameters will be added.
' @return {WebRequest} The WebRequest object with the selected report parameters added as query string parameters.
'
' This function performs the following steps:
' 1. Initializes and displays the 'SelectReportForm' user form.
' 2. If the user cancels the form, raises an error and displays a message.
' 3. Converts the selected date range from the user form to the required format for the Xero API request.
' 4. Adds the 'fromDate' and 'toDate' query string parameters to the WebRequest object with the selected date range.
' 5. Returns the updated WebRequest object.
'
' Note: This function uses the 'TextBox1' and 'TextBox2' controls of the 'SelectReportForm' user form to retrieve the selected date range.
''
Private Function SelectReport(Request As WebRequest) As WebRequest
    On Error GoTo ApiCall_Cleanup

    ' Initialize the 'SelectReportForm' user form
    Dim SelectForm1 As SelectReportForm
    Set SelectForm1 = New SelectReportForm
    
    ' Display the user form
    SelectForm1.show
    
    ' Check if the user canceled the form
    If SelectForm1.UserCancel Then
        ' Notify the user and raise an error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly
        Err.Raise 11040 + vbObjectError, "SelectReportForm", "User canceled selection form"
    End If
    
    ' Convert the selected date range to the required format
    Dim fromDate As Date
    Dim toDate As Date
    fromDate = DateSerial(CInt(Right(SelectForm1.TextBox1.value, 4)), CInt(Left(SelectForm1.TextBox1.value, 2)), CInt(Mid(SelectForm1.TextBox1.value, 4, 2)))
    toDate = DateSerial(CInt(Right(SelectForm1.TextBox2.value, 4)), CInt(Left(SelectForm1.TextBox2.value, 2)), CInt(Mid(SelectForm1.TextBox2.value, 4, 2)))
    
    ' Add the 'fromDate' and 'toDate' query string parameters to the WebRequest object
    Request.AddQuerystringParam "fromDate", Format(fromDate, "yyyy-mm-dd")
    Request.AddQuerystringParam "toDate", Format(toDate, "yyyy-mm-dd")
    
    ' Return the updated WebRequest object
    Set SelectReport = Request

ApiCall_Cleanup:
    ' Unload the user form and handle errors
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

''
' Retrieves a Profit and Loss report from the Xero API for the selected date range.
'
' @method GetPnLReport
' @return {Dictionary} A dictionary containing the Profit and Loss report data, or an empty dictionary if an error occurs.
'
' This function performs the following steps:
' 1. Initializes a new WebRequest object for the API request.
' 2. Configures the WebRequest object with the required parameters for the Profit and Loss report API endpoint.
' 3. Displays the 'SelectReportForm' user form to allow the user to select the report date range.
' 4. Sends the API request to the Xero API using the configured WebRequest object.
' 5. If the API request is successful (200 status code), returns the report data as a dictionary.
' 6. If the API request fails, raises an error with the appropriate error details.
'
' Note: This function uses the 'XeroClient' property to execute the API request and the 'SelectReport' function to obtain the report date range.
''
Private Function GetPnLReport() As Dictionary
    On Error GoTo ApiCall_Cleanup

    ' Initialize a new WebRequest object for the API request
    Dim ReportRequest As WebRequest
    Set ReportRequest = New WebRequest
    
    ' Configure the WebRequest object for the Profit and Loss report API endpoint
    ReportRequest.Resource = "api.xro/2.0/Reports/ProfitAndLoss"
    ReportRequest.Method = WebMethod.HttpGet
    ReportRequest.RequestFormat = WebFormat.FormUrlEncoded
    ReportRequest.ResponseFormat = WebFormat.Json
    
    ' Display the 'SelectReportForm' user form to obtain the report date range
    Set ReportRequest = SelectReport(ReportRequest)
    
    ' Send the API request and retrieve the response
    Dim ReportResponse As WebResponse
    Set ReportResponse = XeroClient.Execute(ReportRequest)
    
    ' If the API request is successful, return the report data
    If ReportResponse.StatusCode = WebStatusCode.Ok Then
        Set GetPnLReport = ReportResponse.Data
    Else
        ' If the API request fails, raise an error
        Err.Raise 11041 + vbObjectError, "XeroAPICall.GetPnLReport", _
            ReportResponse.StatusCode & ": " & ReportResponse.Content
    End If
    
ApiCall_Cleanup:
    ' Clean up objects and handle errors
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

''
' Parses response data (JSON in the form of a Dictionary) and loads it into an Excel sheet.
'
' @method LoadReportToSheet
' @param {Dictionary} GetReportData - The JSON object obtained from an API call.
' @param {String} [SheetName=ReportOutputSheet] - Optional name for the output sheet.
'
' This function performs the following steps:
' 1. Extracts the report and its components from the JSON response.
' 2. Initializes the Excel sheet and sets the starting row index.
' 3. Writes the report titles to the sheet, formatting them appropriately.
' 4. Adds the account and date headers, formatting the date header based on the report period.
' 5. Iterates over the sections and rows of the report, adding data to the sheet and applying styles.
' 6. Adjusts the sheet name to avoid duplicates and applies final formatting.
' 7. Handles any errors that occur and logs them.
''
Private Sub LoadReportToSheet(GetReportData As Dictionary, Optional SheetName As String = ReportOutputSheet)
    On Error GoTo ApiCall_Cleanup

    ' Extract the report data from the JSON object
    Dim report  As Dictionary
    Set report = GetReportData("Reports")(1)
    
    Dim reportTitles As Collection
    Set reportTitles = report("ReportTitles")
    
    Dim rows As Collection
    Set rows = report("Rows")
    
    ' Initialize the row index for the Excel sheet
    Dim rowIndex As Long
    rowIndex = 1 ' Start at the first row

    ' Create a new sheet to load the JSON data into
    Dim sh As Worksheet
    Dim sheetIndex As Integer
    sheetIndex = 0
    
    ' Create new sheets
    Set sh = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    
    ' Write the report titles to the sheet
    With sh
        Dim reportTitle As Variant
        For Each reportTitle In reportTitles
            ' Add titles to the sheet
            If rowIndex = 3 Then
                ' Add period information for the third row title
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
        ' Add a blank row after the titles
        rowIndex = rowIndex + 1

        ' Add the account and date headers
        .Cells(rowIndex, 1).value = "Account"
        
        ' Re-format the report dates
        Dim dates() As String
        Dim dateFormat As String
        dates = Split(reportTitles(3), " to ")
        dateFormat = "d mmm"
        If Year(CDate(dates(0))) <> Year(CDate(dates(1))) Then
            dateFormat = dateFormat & " yyyy"
        End If
        .Cells(rowIndex, 2).value = Format(CDate(dates(0)), dateFormat) & "-" & Format(CDate(dates(1)), "d mmm yyyy")
        SheetName = SheetName & UCase(Format(CDate(dates(1)), "dmmmyy"))
        
        ' Format the header cells
        With .Range(.Cells(rowIndex, 1), .Cells(rowIndex, 2))
            .Font.Bold = True
            .Font.Size = 10
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        ' Add a blank row after the headers
        rowIndex = rowIndex + 2
        
        ' Iterate over the sections and rows of the report
        Dim section As Dictionary
        Dim innerRows As Collection
        Dim row As Dictionary
        
        For Each section In rows
            If section("RowType") = "Section" Then
                If section("Title") <> "" Then
                    .Cells(rowIndex, 1).value = section("Title")
                    ' Style the section header
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
                
                    ' Style each row
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
        ' Set the font for the entire sheet
        .Range(.Cells(1, 1), .Cells(rowIndex, 2)).Font.name = "Arial"
        ' Autofit columns for better layout
        .Range(.Cells(5, 1), .Cells(rowIndex, 2)).Columns.AutoFit
    End With
    
    ' Check for available sheet names to avoid duplicates
    If WebHelpers.WorksheetExists(SheetName, ThisWorkbook) Then
        Do While WebHelpers.WorksheetExists(SheetName & "_" & sheetIndex, ThisWorkbook)
            sheetIndex = sheetIndex + 1
        Loop
        SheetName = SheetName & "_" & sheetIndex
    End If
    sh.name = SheetName
    ' Turn off gridlines for better presentation
    WebHelpers.TurnOffGridLines sh
    
    ' Save the name of the sheet
    pLastOutputSheetName = SheetName
    
ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while loading report to sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.LoadReportToSheet", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAPICall.LoadReportToSheet", auth_ErrorDescription
    End If
End Sub

' --------------------------------------------- '
' Execution
' --------------------------------------------- '

''
' Calls the login procedures for the user interface button.
'
' @method Login_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Retrieves the pre-set authenticator object from the XeroClient.
' 3. Logs out and clears the cache for the current session.
' 4. Initiates the login process.
' 5. Returns the authenticator reference to the XeroClient.
' 6. Handles any errors that occur during the process and logs them.
''
Public Sub Login_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve the pre-set authenticator object
    Dim Auth As XeroAuthenticator
    Set Auth = XeroClient.Authenticator
    Set XeroClient.Authenticator = Nothing
    
    ' Logout and clear cache for the current session
    Auth.Logout
    
    ' Login
    Auth.Login
    
    ' Return the authenticator reference to the XeroClient
    Set XeroClient.Authenticator = Auth
    ' Clear the local reference to the authenticator
    Set Auth = Nothing
    
ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error happened
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the login process." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.Login_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub

''
' Calls the report generation procedures for the user interface button.
'
' @method GenerateReport_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Retrieves the Profit and Loss report from the API.
' 3. Parses and loads the report data into an Excel sheet.
' 4. Displays a message box to notify the user of the successful report generation.
' 5. Handles any errors that occur during the process and logs them.
''
Public Sub GenerateReport_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve the Profit and Loss report from the API
    Dim ReportDict As Dictionary
    Set ReportDict = GetPnLReport
        
    ' Parse and load the report data into an Excel sheet
    LoadReportToSheet ReportDict
    ' Notify the user of successful report generation
    MsgBox "Report successfully generated on sheet: " & vbNewLine & pLastOutputSheetName, vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error occurred
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while generating report." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.GenerateReport_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub

''
' Clears all saved tokens and Xero organizations/tenants ID for the user interface button.
'
' @method ClearCache_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Confirms the user's action to clear the cache.
' 3. If the user confirms, retrieves the pre-set authenticator object.
' 4. Clears all cache (tenants and tokens) and logs out of the current session.
' 5. Returns the authenticator reference to the XeroClient.
' 6. Handles any errors that occur during the process and logs them.
''
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
            ' Retrieve the pre-set authenticator object
            Dim Auth As XeroAuthenticator
            Set Auth = XeroClient.Authenticator
            ' Clear the reference to the authenticator in the XeroClient
            Set XeroClient.Authenticator = Nothing
            
            ' Clear all cache (tenants and tokens)
            Auth.ClearAllCache isClearTenant:=True, isClearToken:=True
            
            ' Clear current session tokens cache by logging out
            Auth.Logout
            
            ' Return the authenticator reference to the XeroClient
            Set XeroClient.Authenticator = Auth
            ' Clear the local reference to the authenticator
            Set Auth = Nothing
            
        Case vbNo
            ' Exit the subroutine if the user cancels the action
            Exit Sub
    End Select

ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error occurred
        pXeroClientId = ""
        pXeroClientSecret = ""
        Set XeroClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAPICall.ClearCache_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
    End If
End Sub
