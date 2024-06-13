VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XeroAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' XeroAPICall v1.0.0
' @author musayohanes00@gmail.com
' https://github.com/Muyoouu/vba-xero-api
'
' Xero Authenticator
' Custom IWebAuthenticator for "Xero" authentication for Xero Accounting API
' Source (IWebAuthenticator) credits: Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Xero authentication docs:
' - https://developer.xero.com/documentation/guides/oauth2/overview
'
' Other docs:
' - Register for Client Id and Client Secret: https://developer.xero.com/app/manage/
' - List of available scopes: https://developer.xero.com/documentation/guides/oauth2/scopes
'
' @example
' ```VB.net
' Dim Auth As XeroAuthenticator
' Set Auth = New XeroAuthenticator
' Auth.Setup "Your Client Id", "Your Client Secret"
'
' ' Add scopes, for example request for accounting reports read
' ' Request refresh token with 'offline_access' scope included
' Auth.AddScope "offline_access"
' Auth.AddScope "accounting.reports.read"
'
' ' Manually open up Xero login
' ' (called automatically on first request otherwise)
' Auth.Login
'
' ' Add authenticator to client
' Set Client.Authenticator = Auth
' ```
'
' @class XeroAuthenticator
' @implements IWebAuthenticator v4.*
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' URL for login process
Private Const auth_AuthorizationUrl As String = "https://login.xero.com/identity/connect/authorize"
' Redirect URL registered inside Xero
Private Const auth_RedirectUrl As String = "https://developer.xero.com/"
' Sheet name to store reusable tokens (refresh token - valid for 60 days)
Private Const auth_SheetTokensStorage As String = "Sheet_XeroTokens"
' Table(ListObject) name to store Xero tenants info (and ID)
Private Const auth_TableTokensStorage As String = "XeroTokens"
' Sheet name to store Xero tenants info (and ID)
Private Const auth_SheetTenantsStorage As String = "Sheet_XeroTenants"
' Table(ListObject) name to store Xero tenants info (and ID)
Private Const auth_TableTenantsStorage As String = "XeroTenants"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public ClientSecret As String
Public ApiKey As String
Public AuthorizationCode As String
Public AccessToken As String
Public AccessTokenExpiration As Date
Public RefreshToken As String
Public RefreshTokenExpiration As Date
Public Scopes As Variant
Private pDTenantsId As Dictionary

''
' A dictionary that stores the Xero organizations/tenants IDs.
' The keys of the dictionary represent the organization/tenant names, and the values represent the corresponding tenant IDs.
'
' @property DTenantsId
' @type {Dictionary}
''
Public Property Get DTenantsId() As Dictionary
    ' Initialize the dictionary if it doesn't exist
    If pDTenantsId Is Nothing Then
        Set pDTenantsId = New Dictionary
        ' Set the comparison mode to text-based comparison
        pDTenantsId.CompareMode = TextCompare
    End If
    
    ' Return the dictionary reference
    Set DTenantsId = pDTenantsId
End Property


' ============================================= '
' Public Methods
' ============================================= '

''
' Sets up the authenticator properties required for Xero API authentication.
'
' @method Setup
' @param {String} ClientId - The client ID provided by Xero for your application.
' @param {String} ClientSecret - The client secret provided by Xero for your application.
'
' This method initializes the 'ClientId' and 'ClientSecret' properties with the provided values.
' These properties are used during the authentication and token exchange processes.
''
Public Sub Setup(ClientId As String, ClientSecret As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
End Sub

''
' Direct user to login and authenticate in Xero website via browser.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#1-send-a-user-to-authorize-your-app
'
' @method Login
'
' The login process involves the following steps:
' 1. Open a browser instance using the ChromeDevProtocols framework (https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA).
' 2. Wait for the user to complete the login process, either by approval or encountering an error.
' 3. Extract the authentication code from the browser URL if the login is approved.
' 4. Handle any errors that occur during the login process.
'
' If the login is successful, the 'AuthorizationCode' property is set with the extracted authentication code.
' If an error occurs or the login is canceled by the user, an appropriate error message is displayed.
''
Public Sub Login()
    On Error GoTo auth_Cleanup
    
    ' Open browser using ChromeDevProtocols framework
    Dim objBrowser As New CDPBrowser
    objBrowser.start name:="edge", appUrl:=Me.GetLoginUrl, cleanActive:=True, reAttach:=True
    
    ' Boolean values to track login completion
    Dim auth_IsApprove As Boolean
    Dim auth_IsError As Boolean
    
    ' Wait for login to complete, either by approval or error
    Do
        auth_IsApprove = auth_LoginIsApproval(objBrowser)
        auth_IsError = auth_LoginIsError(objBrowser)
        ' Sleep to avoid overloading the system
        objBrowser.sleep 1
    Loop Until auth_IsApprove Or auth_IsError
    
    ' Handle errors or extract the authentication code
    If auth_IsError Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", auth_LoginExtractError(objBrowser)
    Else
        Me.AuthorizationCode = auth_LoginExtractCode(objBrowser)
        If Me.AuthorizationCode = "" Then
            MsgBox "Login might be canceled! process is stopped", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was canceled"
        End If
    End If
    
auth_Cleanup:
    ' Error handling and cleanup
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        auth_ErrorDescription = "An error occurred while logging in." & vbNewLine
        If Err.Number - vbObjectError <> 11040 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Close browser
        If Not objBrowser Is Nothing Then: objBrowser.quit
        Set objBrowser = Nothing
        
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.Login", 11040 + vbObjectError
        Err.Raise 11040 + vbObjectError, "XeroAuthenticator.Login", auth_ErrorDescription
    Else
        ' Close browser
        If Not objBrowser Is Nothing Then: objBrowser.quit
        Set objBrowser = Nothing
        
        If Not auth_IsApprove Then
            auth_ErrorDescription = "Login did not complete." & vbNewLine
            
            WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.Login", 11040 + vbObjectError
            Err.Raise 11040 + vbObjectError, "XeroAuthenticator.Login", auth_ErrorDescription
        Else
            WebHelpers.LogDebug "Login succeeded - auth code: " & Me.AuthorizationCode, "XeroAuthenticator.Login"
            ' Inform user that the login process succeeded
            MsgBox "Xero login succeeded!", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        End If
    End If
End Sub

''
' Logs out by clearing all cached tokens and authorizations, effectively resetting the login session.
' After calling this method, the user will need to log in again to obtain new tokens and authorizations.
'
' @method Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.AccessToken = ""
    Me.AccessTokenExpiration = CDate(0)
    Me.RefreshToken = ""
    Me.RefreshTokenExpiration = CDate(0)
End Sub

''
' Adds a new scope to the list of authorized scopes for the Xero API.
' To get a list of available scopes, visit https://developer.xero.com/documentation/guides/oauth2/scopes/
' The 'Scopes' property is used during the authentication process to request access to specific areas of the Xero API.
'
' @method AddScope
' @param {String} Scope - The scope to add to the authorized scopes list.
''
Public Sub AddScope(Scope As String)
    Dim auth_Scopes As Variant
    
    ' Retrieve the current list of authorized scopes
    auth_Scopes = Me.Scopes
    
    ' Initialize a new array if the 'Scopes' property is empty
    If VBA.IsEmpty(auth_Scopes) Then
        ReDim auth_Scopes(0 To 0)
    Else
        ' Resize the array to accommodate the new scope
        ReDim Preserve auth_Scopes(0 To UBound(auth_Scopes) + 1)
    End If
    
    ' Add the new scope to the end of the array
    auth_Scopes(UBound(auth_Scopes)) = Scope
    
    ' Update the 'Scopes' property with the modified array
    Me.Scopes = auth_Scopes
End Sub

''
' Obtains a new or refreshed access token and refresh token, if necessary.
'
' @method GetToken
' @param {WebClient} Client - The WebClient instance to use for token requests.
'
' This method follows the following logic:
' 1. If the current access token exists and is not expired, no action is taken.
' 2. If the current refresh token exists and is not expired, the 'GetRefreshedToken' method is called to obtain a new access token and refresh token.
' 3. If no valid access token or refresh token exists:
'    a. If the 'AuthorizationCode' is empty, the 'Login' method is called to obtain a new authorization code.
'    b. The 'GetNewToken' method is called to obtain a new access token and refresh token using the authorization code.
' 4. The new or refreshed tokens are stored by calling the 'StoreTokens' method.
'
' Note: This method uses and updates the 'AccessToken', 'AccessTokenExpiration', 'RefreshToken', 'RefreshTokenExpiration', and 'AuthorizationCode' properties.
''
Public Sub GetToken(Client As WebClient)
        
    ' Check if the current access token exists and is not expired
    If Me.AccessToken <> "" And Not Me.AccessTokenExpiration < Now Then
        WebHelpers.LogDebug "AccessToken already exist and not expired, skipping GetToken process", "XeroAuthenticator.GetToken"
        Exit Sub
    ' Check if the current refresh token exists and is not expired
    ElseIf Me.RefreshToken <> "" And Not Me.RefreshTokenExpiration < Now Then
        WebHelpers.LogDebug "RefreshToken exist and not expired, attempting to refresh AccessToken", "XeroAuthenticator.GetToken"
        GetRefreshedToken Client
    ' If no valid tokens exist, obtain a new access token and refresh token
    Else
        ' Login first if the authorization code doesn't exist
        If Me.AuthorizationCode = "" Then
            WebHelpers.printMsg "XeroAuthenticator.GetToken", "Auth code not found, redirecting to Login"
            Me.Login
        End If
        WebHelpers.LogDebug "Tokens not found or expired, getting new tokens", "XeroAuthenticator.GetToken"
        GetNewToken Client
    End If
    
    ' Store the new or refreshed tokens
    StoreTokens
End Sub

''
' Requests the list of authorized Xero organizations/tenants from the Xero API.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#5-check-the-tenants-youre-authorized-to-access
'
' @method GetTenant
' @param {WebClient} Client - The WebClient instance to use for the API request.
'
' This method performs the following steps:
' 1. Clear the 'DTenantsId' dictionary to store the new list of tenants.
' 2. Prepare the API request to retrieve the list of authorized tenants.
' 3. Execute the API request using the provided WebClient instance.
' 4. If the request is successful (200 status code), store the tenant information in the 'DTenantsId' dictionary and on an Excel sheet.
' 5. If the response is empty (no connected organizations), display a message and raise an error.
' 6. Handle any other errors that occur during the API request.
'
' Note: This method uses the 'AccessToken' property and updates the 'DTenantsId' dictionary.
''
Public Sub GetTenant(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TenantClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clear the 'DTenantsId' dictionary
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TenantClient = Client.Clone
    Set auth_TenantClient.Authenticator = Nothing
    auth_TenantClient.BaseUrl = "https://api.xero.com/"
    
    ' Prepare the API request to retrieve the list of authorized tenants
    auth_Request.Resource = "connections"
    auth_Request.Method = WebMethod.HttpGet
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    Set auth_Response = auth_TenantClient.Execute(auth_Request)
    
    ' Process the API response
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        If auth_Response.Data.Count > 0 Then
            ' Load response into "DTenantsId" property
            Dim TenantDict As Dictionary
            For Each TenantDict In auth_Response.Data
                DTenantsId(CStr(TenantDict("tenantName"))) = CStr(TenantDict("tenantId"))
            Next TenantDict
            
            ' Store response into excel sheet
            StoreTenants auth_Response
            
            WebHelpers.LogDebug "GetTenant success, available tenants: " & CStr(DTenantsId.Count), "XeroAuthenticator.GetTenant"
        Else
            MsgBox "Cannot find any Xero organization connected to your account, try to re-login and authorize a connection!", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", _
                auth_Response.StatusCode & ": " & auth_Response.Content & vbNewLine & _
                "Empty response; no connected organization, need to authorize and connect through Login"
        End If
    Else
        ' Raise an error if the API request is not successful
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TenantClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving new tenants." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.GetTenant", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", auth_ErrorDescription
    End If
End Sub

''
' Displays a user form that allows the user to select the Xero organization/tenant for the API request.
'
' @method SelectTenant
' @param {WebClient} Client - The WebClient instance to use for retrieving the list of tenants if necessary.
' @return {String} The selected tenant ID, or an empty string if the user cancels the selection.
'
' This function performs the following steps:
' 1. Loads the list of tenants from the saved Excel sheet using the 'LoadTenants' method.
' 2. If no tenants are loaded, calls the 'GetTenant' method to retrieve the list of tenants from the Xero API.
' 3. Initializes and displays the 'SelectXeroOrgForm' user form with the list of tenants.
' 4. If the user requests an update, calls the 'GetTenant' method again and refreshes the user form with the updated list.
' 5. If the user cancels the selection, raises an error and displays a message.
' 6. Returns the selected tenant ID if a tenant is selected.
'
' Note: This function uses and updates the 'DTenantsId' property.
''
Public Function SelectTenant(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    ' Load the list of tenants from the saved Excel sheet
    LoadTenants
    
    ' If no tenants are loaded, retrieve the list from the Xero API
    If Not DTenantsId.Count > 0 Then
        GetTenant Client
    End If

    ' Initialize the 'SelectXeroOrgForm' user form
    Dim SelectForm1 As SelectXeroOrgForm
    Set SelectForm1 = New SelectXeroOrgForm
    
    ' Assign the list of tenants to the user form
    SelectForm1.OrgList = DTenantsId.Keys
    
    ' Display the user form
    SelectForm1.show
    
    ' Handle user requests for updating the tenant list
    If SelectForm1.UserRequestUpdate Then
        ' Get and update the tenant list
        GetTenant Client
        SelectForm1.OrgList = DTenantsId.Keys
        
        ' Reset the form property to false
        SelectForm1.UserRequestUpdate = False
        
        ' Show the form again with the updated list
        SelectForm1.show
    End If
    
    ' Check if the user canceled the selection
    If SelectForm1.UserCancel Then
        ' Notify the user and raise an error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Err.Raise 11040 + vbObjectError, "SelectXeroOrgForm", "User canceled selection form"
    End If
    
    ' Return the selected tenant ID
    SelectTenant = DTenantsId(CStr(SelectForm1.ComboBox1.value))

auth_Cleanup:
    ' Unload the user form and handle errors
    If Not SelectForm1 Is Nothing Then
        Unload SelectForm1
    End If
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the user selection process for a Xero organization." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.SelectTenant", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.SelectTenant", auth_ErrorDescription
    End If
End Function

''
' Clears/deletes all saved cache (tenants & tokens) from this Excel workbook.
'
' @method ClearAllCache
'
' This function performs the following steps:
' 1. Checks the boolean flags to determine which cache to clear (tenants, tokens, or both).
' 2. If 'isClearTenant' is True, it calls the 'ClearTenants' function to clear tenant information.
' 3. If 'isClearToken' is True, it calls the 'ClearTokens' function to clear token information.
' 4. Handles any errors that occur during the process and logs them.
''
Public Sub ClearAllCache(isClearTenant As Boolean, isClearToken As Boolean)
    On Error GoTo auth_Cleanup
    
    ' Check if tenant cache needs to be cleared
    If isClearTenant Then
        ' Call the ClearTenants function
        ClearTenants
    End If
    
    ' Check if token cache needs to be cleared
    If isClearToken Then
        ' Call the ClearTokens function
        ClearTokens
    End If
    
auth_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearAllCache", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearAllCache", auth_ErrorDescription
    End If
End Sub

''
' Hook for taking action before a request is executed; used for setting up required request headers for API calls.
' This method is called before every API request made to the Xero API.
'
' Xero official documentation on required headers for API calls:
' https://developer.xero.com/documentation/guides/oauth2/auth-flow/#6-call-the-api
'
' @param {WebClient} Client - The WebClient instance that is about to execute the request.
' @param in|out {WebRequest} Request - The request object that is about to be executed.
'
' This method performs the following steps:
' 1. Call the 'GetToken' method to ensure that the current access token is valid and not expired.
' 2. Set the 'Authorization' header with the current access token.
' 3. Call the 'SelectTenant' method to allow the user to select the Xero organization/tenant for the API request.
' 4. Set the 'xero-tenant-id' header with the selected tenant ID.
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' Ensure a valid access token is available for the API request
    Me.GetToken Client
    
    ' Set the 'Authorization' header with the current access token
    Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    ' Allow the user to select the Xero organization/tenant for the API request
    Dim SelectedTenantId As String
    SelectedTenantId = SelectTenant(Client)
    
    ' Set the 'xero-tenant-id' header with the selected tenant ID
    Request.SetHeader "xero-tenant-id", SelectedTenantId
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
End Sub

''
' Formats the URL for the Xero login process, including the required scopes and other parameters.
'
' @method GetLoginUrl
' @return {String} The formatted login URL.
'
' The formatted URL includes the following parameters:
' - response_type: Set to "code" to indicate that an authorization code will be returned.
' - client_id: The client ID provided by Xero for authentication.
' - redirect_uri: The predetermined redirect URL to which the user will be directed after authentication.
' - state: An optional state parameter, currently set to a hardcoded value of "123".
' - scope: The list of scopes (permissions) requested for the application. If no scopes are specified, an empty string is used.
'
' Docs: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#1-send-a-user-to-authorize-your-app
''
Public Function GetLoginUrl() As String
    ' Use WebRequest object as helper to format URL
    Dim auth_Request As New WebRequest
    auth_Request.Resource = auth_AuthorizationUrl
    
    ' Add required parameters
    auth_Request.AddQuerystringParam "response_type", "code"
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "state", "123"
    
    ' Add scopes, or an empty string if no scopes are specified
    If Not VBA.IsEmpty(Me.Scopes) Then
        auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, " ")
    Else
        auth_Request.AddQuerystringParam "scope", ""
    End If
    
    ' Get the formatted URL and clean up
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

''
' Checks if the current URL of the browser matches the predetermined redirect URL stored in the 'auth_RedirectUrl' constant.
'
' @method auth_LoginIsApproval
' @param {CDPBrowser} objBrowser - The browser object to check the URL from.
' @return {Boolean} True if the current URL matches the redirect URL, False otherwise.
''
Private Function auth_LoginIsApproval(objBrowser As CDPBrowser) As Boolean
    Dim auth_BrowserUrlParts As Dictionary
    Dim auth_RedirectUrlParts As Dictionary
    
    Set auth_BrowserUrlParts = WebHelpers.GetUrlParts(objBrowser.url)
    Set auth_RedirectUrlParts = WebHelpers.GetUrlParts(auth_RedirectUrl)
    
    ' Check if the 'Host' part of both URLs match.
    If auth_BrowserUrlParts("Host") = auth_RedirectUrlParts("Host") Then
        auth_LoginIsApproval = True
    Else
        auth_LoginIsApproval = False
    End If
End Function

''
' Checks if the current HTML page rendered in the browser contains an error message in the title.
'
' @method auth_LoginIsError
' @param {CDPBrowser} objBrowser - The browser object to check for errors.
' @return {Boolean} True if the current page title contains the word 'error', False otherwise.
''
Private Function auth_LoginIsError(objBrowser As CDPBrowser) As Boolean
    Dim auth_Element As CDPElement
    
    ' Wait for the page to fully load before checking for errors.
    objBrowser.wait till:=ReadyState.isComplete
    
    ' Search for the 'error' word in the HTML title.
    Set auth_Element = objBrowser.getElementByQuery("title", True)
    
    ' If the title element exists, check if it contains the word 'error'.
    If auth_Element.isExist Then
        auth_LoginIsError = InStr(1, CStr(auth_Element.innerText), "error", vbTextCompare) > 0
    Else
        ' If the title element does not exist, assume no error.
        auth_LoginIsError = False
    End If
End Function

''
' Extracts the authentication code from the current URL of the browser if it matches the predetermined redirect URL.
'
' @method auth_LoginExtractCode
' @param {CDPBrowser} objBrowser - The browser object to extract the URL from.
' @return {String} The extracted authentication code, or an empty string if the URL does not match the redirect URL.
''
Private Function auth_LoginExtractCode(objBrowser As CDPBrowser) As String
    ' Check if the current URL matches the predetermined redirect URL.
    If auth_LoginIsApproval(objBrowser) Then
        Dim auth_UrlParts As Dictionary
        Dim auth_ParsedUrlParams As Dictionary
        
        ' Extract the authorization code from the 'code' parameter in the URL query string.
        Set auth_UrlParts = WebHelpers.GetUrlParts(objBrowser.url)
        Set auth_ParsedUrlParams = ParseUrlEncoded(auth_UrlParts("Querystring"))
        auth_LoginExtractCode = auth_ParsedUrlParams("code")
    Else
        ' If the URL does not match the redirect URL, return an empty string.
        auth_LoginExtractCode = ""
    End If
End Function

''
' Extracts the error descriptions in the current HTML page rendered in the browser.
'
' @method auth_LoginExtractError
' @param {CDPBrowser} objBrowser - The browser object to extract error descriptions from.
' @return {String} Returns extracted error descriptions if found, otherwise "Unknown".
'
' This function performs the following steps:
' 1. Uses the browser object to find an element containing the text "Error".
' 2. If the error element is found, it extracts the inner text of the element as the error code.
' 3. Then, it looks for a following sibling element of the error element to extract the error description.
' 4. If both elements are found, it concatenates the error code and description.
' 5. If no error elements are found, it returns "Unknown".
''
Private Function auth_LoginExtractError(objBrowser As CDPBrowser) As String
    Dim auth_Element As CDPElement
    
    ' Attempt to find an element containing "Error" in its text
    Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]")
    
    If auth_Element.isExist Then
        ' Extract the error code from the element
        auth_LoginExtractError = auth_Element.innerText
        
        ' Attempt to find the following sibling element containing the error description
        Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]/following-sibling::p")
        If auth_Element.isExist Then
            ' Concatenate the error code and the error description
            auth_LoginExtractError = auth_LoginExtractError & " - " & auth_Element.innerText
        End If
    Else
        ' Return "Unknown" if no error element is found
        auth_LoginExtractError = "Unknown"
    End If
End Function

''
' Exchanges the current AuthorizationCode for new access and refresh tokens.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#3-exchange-the-code
'
' @method GetNewToken
' @param {WebClient} Client - The WebClient instance to use for the token exchange request.
'
' This method performs the following steps:
' 1. Prepare the token request by setting the required headers, request body, and other parameters.
' 2. Execute the token request using the provided WebClient instance.
' 3. If the request is unauthorized (401 status code), attempt to re-login and update the AuthorizationCode, then retry the token request.
' 4. If the token request is successful (200 status code), update the AccessToken, AccessTokenExpiration, RefreshToken, and RefreshTokenExpiration properties.
' 5. Handle any errors that occur during the token exchange process.
'
' Note: This method uses the 'ClientId', 'ClientSecret', 'AuthorizationCode', and 'auth_RedirectUrl' properties.
''
Private Sub GetNewToken(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://identity.xero.com/"
    
    ' Prepare the token request
    auth_Request.Resource = "connect/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "authorization", "Basic " & Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare the request body
    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "redirect_uri", auth_RedirectUrl
    auth_Body.Add "grant_type", "authorization_code"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Response 401 - auth code might be expired
    ' Handle unauthorized (401) response by attempting to re-login and update the AuthorizationCode
    If auth_Response.StatusCode = WebStatusCode.Unauthorized Then
        ' Log the response details
        WebHelpers.printMsg "XeroAuthenticator.GetNewToken", auth_Response.StatusDescription & " - " & auth_Response.Data("details")
        
        ' Notify and ask permission to attempt re-login
        Dim msgBoxResponse As VbMsgBoxResult
        msgBoxResponse = MsgBox("Login session expired:" & vbNewLine & auth_Response.Data("details") & _
            vbNewLine & vbNewLine & "Attempt to re-login?", vbQuestion + vbYesNo, "Xero Report Generator - Microsoft Excel")
        
        Select Case msgBoxResponse
            Case vbYes
                WebHelpers.printMsg "XeroAuthenticator.GetNewToken", "Redirecting to Login..."
        
                ' Clear expired auth code and token
                Me.Logout
                
                ' Re-login and update the auth code
                Me.Login
                auth_Request.Body("code") = Me.AuthorizationCode
                
                ' Re-attempt request to get token
                Set auth_Response = auth_TokenClient.Execute(auth_Request)
            
            ' Notify stopping the get token process
            ' Leaving response status at 401 will raise error eventually
            Case vbNo
                WebHelpers.printMsg "XeroAuthenticator.GetNewToken", "Abort get token process"
        End Select
    End If
    
    ' Update token properties if the token request is successful
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        Me.AccessToken = auth_Response.Data("access_token")
        Me.AccessTokenExpiration = DateAdd("s", CLng(auth_Response.Data("expires_in")), Now)
        If auth_Response.Data.Exists("refresh_token") Then
            Me.RefreshToken = auth_Response.Data("refresh_token")
            Me.RefreshTokenExpiration = DateAdd("d", 60, Now)
        End If
        
        WebHelpers.LogDebug "New Access Token: " & Me.AccessToken & vbNewLine & vbNewLine & _
            "New Refresh Token: " & vbNewLine & Me.RefreshToken & vbNewLine, "XeroAuthenticator.GetNewToken"
    Else
        ' Raise an error if the token request is not successful
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetNewToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving new token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.GetNewToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetNewToken", auth_ErrorDescription
    End If
End Sub

''
' Exchanges the current refresh token for new access and refresh tokens.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#refreshing-access-and-refresh-tokens
'
' @method GetRefreshedToken
' @param {WebClient} Client - The WebClient instance to use for the token exchange request.
'
' This method performs the following steps:
' 1. Prepare the token refresh request by setting the required headers, request body, and other parameters.
' 2. Execute the token refresh request using the provided WebClient instance.
' 3. If the request is successful (200 status code), update the AccessToken, AccessTokenExpiration, RefreshToken, and RefreshTokenExpiration properties.
' 4. If the request is unauthorized (401 status code), attempt to re-login and update the AuthorizationCode, then redirect to the GetNewToken procedure.
' 5. Handle any other errors that occur during the token refresh process.
'
' Note: This method uses the 'ClientId', 'ClientSecret', and 'RefreshToken' properties.
''
Private Sub GetRefreshedToken(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://identity.xero.com/"
    
    ' Prepare the token refresh request
    auth_Request.Resource = "connect/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "authorization", "Basic " + Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare the request body
    auth_Body.Add "refresh_token", Me.RefreshToken
    auth_Body.Add "grant_type", "refresh_token"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Update token properties if the token refresh request is successful
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        Me.AccessToken = auth_Response.Data("access_token")
        Me.AccessTokenExpiration = DateAdd("s", CLng(auth_Response.Data("expires_in")), Now)
        Me.RefreshToken = auth_Response.Data("refresh_token")
        Me.RefreshTokenExpiration = DateAdd("d", 60, Now)
        
        WebHelpers.LogDebug "New Access Token: " & Me.AccessToken & vbNewLine & vbTab & _
            "New Refresh Token: " & Me.RefreshToken, "XeroAuthenticator.GetNewToken"
    
    ' Response 401 - refresh token might be expired
    ' Handle unauthorized (401) response by attempting to re-login and update the AuthorizationCode
    ElseIf auth_Response.StatusCode = WebStatusCode.Unauthorized Then
        ' Log the response details
        WebHelpers.printMsg "XeroAuthenticator.GetRefreshedToken", auth_Response.StatusDescription & " - " & auth_Response.Data("details")
        WebHelpers.LogDebug "XeroAuthenticator.GetRefreshedToken", auth_Response.StatusCode & ": " & auth_Response.Content
        
        ' Notify and ask permission to attempt re-login
        Dim msgBoxResponse As VbMsgBoxResult
        msgBoxResponse = MsgBox("Login session expired:" & vbNewLine & auth_Response.Data("details") & _
            vbNewLine & vbNewLine & "Attempt to re-login?", vbQuestion + vbYesNo, "Xero Report Generator - Microsoft Excel")
        
        Select Case msgBoxResponse
            Case vbYes
                WebHelpers.printMsg "XeroAuthenticator.GetRefreshedToken", "Redirecting to Login..."
        
                ' Clear expired auth code and token
                Me.Logout
                
                ' Re-login and update the auth code
                Me.Login
                
                ' Clean up and redirect to the GetNewToken procedure
                Set auth_TokenClient = Nothing
                Set auth_Request = Nothing
                Set auth_Response = Nothing
                
                GetNewToken Client
            Case vbNo
                ' Notify stopping refresh token process
                WebHelpers.printMsg "XeroAuthenticator.GetRefreshedToken", "Abort refreshing token process"
        End Select
    Else
        ' Raise an error if the token refresh request is not successful
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetRefreshedToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.GetNewToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetNewToken", auth_ErrorDescription
    End If
End Sub

''
' Saves obtained access token and refresh token into an Excel sheet.
'
' @method StoreTokens
'
' This function performs the following steps:
' 1. Checks if the sheet for storing tokens exists. If not, creates a new sheet and hides it.
' 2. Checks if the table (ListObject) for storing tokens exists. If not, creates a new table with headers.
' 3. Adds a new row to the table and fills it with the token data.
' 4. Limits the number of rows in the table to 20, deleting any excess rows from the bottom.
' 5. Logs the success of the operation and handles any errors that occur.
''
Private Sub StoreTokens()
    On Error GoTo StoreTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    Dim NewRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Check if the sheet for storing tokens exists
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        ' Create a new sheet and set its name
        Set sh = wb.Sheets.Add(Before:=wb.Sheets(1))
        sh.name = auth_SheetTokensStorage
        
        ' Hide the sheet from the user
        sh.Visible = xlSheetVeryHidden
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check if the ListObject (table) for storing tokens exists
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        ' Create the table headers for storing tokens
        sh.Cells(1, 1).value = "Get_Datetime"
        sh.Cells(1, 2).value = "Access_Token"
        sh.Cells(1, 3).value = "AT_Expired_Datetime"
        sh.Cells(1, 4).value = "Refresh_Token"
        sh.Cells(1, 5).value = "RT_Expired_Datetime"
        
        'Set up the new table as an Excel table (ListObject)
        Set TokenTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                        source:=sh.Range("A1:E1"), _
                                        XlListObjectHasHeaders:=xlYes)
        TokenTable.name = auth_TableTokensStorage
    Else
        ' Table exists, set the table variable
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    ' Add a new row to the table and fill it with token data
    Set NewRow = TokenTable.ListRows.Add(Position:=1, AlwaysInsert:=True)
    NewRow.Range = Array(Now, _
                         Me.AccessToken, _
                         Me.AccessTokenExpiration, _
                         Me.RefreshToken, _
                         Me.RefreshTokenExpiration)
                         
    ' Log success of storing the tokens
    WebHelpers.LogDebug "Stored existing tokens success, sheet: " & auth_SheetTokensStorage, "XeroAuthenticator.StoreTokens"
    
    ' Limit the number of rows in the TokenTable to 20, deleting any excess rows from the bottom
    LimitTableRows TokenTable, 20
    
StoreTokens_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    Set NewRow = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while storing tokens into Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.StoreTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.StoreTokens", auth_ErrorDescription
    End If
End Sub

''
' Loads saved access token and refresh token from an Excel sheet.
'
' @method LoadTokens
'
' This function performs the following steps:
' 1. Checks if the sheet for loading tokens exists. If not, logs the failure and exits.
' 2. Checks if the table (ListObject) for loading tokens exists. If not, logs the failure and exits.
' 3. Checks if there are any rows in the table. If not, logs the failure and exits.
' 4. Retrieves the most recent token record from the table.
' 5. Loads the tokens into the auth object.
' 6. Checks if the access token has expired. If so, logs the failure and clears the token values.
' 7. Checks if the refresh token has expired. If so, logs the failure and clears the token values.
' 8. Handles any errors that occur and logs them.
''
Private Sub LoadTokens()
    On Error GoTo LoadTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    Dim NewestRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Check if the sheet for loading tokens exists
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Load tokens failed, sheet: " & auth_SheetTokensStorage & "does not exist", "XeroAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check if the ListObject (table) for loading tokens exists
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Load tokens failed, ListObject(Table): " & auth_TableTokensStorage & " does not exist", "XeroAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    ' Check if there are any rows in the table
    If Not TokenTable.ListRows.Count > 0 Then
        ' Log failure due to no rows in the table
        WebHelpers.LogDebug "Load tokens failed, ListObject(Table) " & auth_TableTokensStorage & " does not have any rows/records", "XeroAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Get the most recent record of tokens from the top row (index = 1)
        Set NewestRow = TokenTable.ListRows(1)
    End If
    
    ' Load tokens into the auth object
    Me.AccessToken = CStr(NewestRow.Range.Cells(1, 2).value)
    Me.AccessTokenExpiration = CDate(NewestRow.Range.Cells(1, 3).value)
    Me.RefreshToken = CStr(NewestRow.Range.Cells(1, 4).value)
    Me.RefreshTokenExpiration = CDate(NewestRow.Range.Cells(1, 5).value)
    
    ' Check if the access token has expired
    If Me.AccessTokenExpiration < Now Then
        ' Log failure due to expired access token
        WebHelpers.LogDebug "Load AccessToken failed, AccessToken expired at: " & Format(Me.AccessTokenExpiration, "dd/mm/yyyy hh:mm"), "XeroAuthenticator.LoadTokens"
        ' Clear the access token values
        Me.AccessToken = ""
        Me.AccessTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load AccessToken success: " & Me.AccessToken, "XeroAuthenticator.LoadTokens"
    End If
    
    ' Check if the refresh token has expired
    If Me.RefreshTokenExpiration < Now Then
        ' Log failure due to expired refresh token
        WebHelpers.LogDebug "Load RefreshToken failed, RefreshToken expired at: " & Format(Me.RefreshTokenExpiration, "dd/mm/yyyy hh:mm"), "XeroAuthenticator.LoadTokens"
        ' Clear the refresh token values
        Me.RefreshToken = ""
        Me.RefreshTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load RefreshToken success: " & Me.RefreshToken, "XeroAuthenticator.LoadTokens"
    End If
        
LoadTokens_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    Set NewestRow = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while loading tokens from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.LoadTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.LoadTokens", auth_ErrorDescription
    End If
End Sub

''
' Limits the number of rows in a table, deleting rows if the count exceeds a specified limit.
' Used for limiting the rows/records in the tokens table within the StoreTokens procedure.
'
' @method LimitTableRows
' @param {ListObject} Table - The table to limit.
' @param {Integer} IRowsLimit - The maximum number of rows allowed.
'
' This function performs the following steps:
' 1. Checks if the row limit and the number of rows in the table are greater than zero.
' 2. If either is zero, logs the issue and exits.
' 3. If the number of rows is already within the limit, logs the status and exits.
' 4. If the number of rows exceeds the limit, enters a loop to delete rows from the bottom until the limit is met.
' 5. Logs the successful completion of the row limiting process.
''
Private Sub LimitTableRows(Table As ListObject, IRowsLimit As Integer)
    
    Dim TblRows As ListRows
    Set TblRows = Table.ListRows
    
    ' Exit sub if the limit or table rows count is zero
    If Not IRowsLimit > 0 Or Not TblRows.Count > 0 Then
        WebHelpers.LogDebug "Table " & Table.name & " - limit number or table rows can not be zero - " & _
                            "IRowsLimit: " & IRowsLimit & _
                            ", TblRows.Count: " & TblRows.Count, _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    ' Exit sub if the table rows count is already under or equal to the limit
    ElseIf Not TblRows.Count > IRowsLimit Then
        WebHelpers.LogDebug "Table " & Table.name & " - rows is already under/equal to the limit", _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    End If
    
    ' Loop while the table rows count is greater than the limit
    Do While TblRows.Count > IRowsLimit
        ' Delete the bottom row from the table
        TblRows(TblRows.Count).Delete
    Loop
    
    ' Log successful row limiting
    WebHelpers.LogDebug "Table " & Table.name & " - successfully limited to only " & TblRows.Count & " rows", _
                        "XeroAuthenticator.LimitTableRows"
End Sub

''
' Saves obtained Xero tenants' information into an Excel sheet.
'
' @method StoreTenants
' @param {WebResponse} GetTenantResponse - The response received from the GetTenant procedure.
'
' This function performs the following steps:
' 1. Checks if the sheet for storing tenant information exists. If not, creates a new sheet and hides it.
' 2. Checks if the table (ListObject) for storing tenant information exists. If not, creates a new table with headers based on the response keys.
' 3. Adds a new row to the table for each tenant and fills it with the tenant data.
' 4. Limits the number of rows in the table to 50, deleting any excess rows from the bottom.
' 5. Logs the success of the operation and handles any errors that occur.
''
Private Sub StoreTenants(GetTenantResponse As WebResponse)
    On Error GoTo StoreTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Check if the sheet for storing tenant information exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Create a new sheet and set its name
        Set sh = wb.Sheets.Add(Before:=wb.Sheets(1))
        sh.name = auth_SheetTenantsStorage
        
        ' Hide the sheet from the user
        sh.Visible = xlSheetVeryHidden
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) for storing tenant information exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Create the table headers for storing tenant information
        ' Expected 'Data(1)' example from get tenants response:
        '   "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1",
        '   "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d",
        '   "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056",
        '   "tenantType": "ORGANISATION",
        '   "tenantName": "Maple Florist",
        '   "createdDateUtc": "2019-07-09T23:40:30.1833130",
        '   "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        '
        ' Based on official docs (10 May 2024):
        ' https://developer.xero.com/documentation/guides/oauth2/auth-flow/#5-check-the-tenants-youre-authorized-to-access
        '
        ' Column to record GetTenant response datetime
        sh.Cells(1, 1).value = "Get_Datetime"
        ' Columns based on response Data Keys
        sh.Range(sh.Cells(1, 2), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)) = GetTenantResponse.Data(1).Keys
        
        ' Set up the new table as an Excel table (ListObject)
        Set TenantTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                            source:=sh.Range(sh.Cells(1, 1), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)), _
                                            XlListObjectHasHeaders:=xlYes)
        TenantTable.name = auth_TableTenantsStorage
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Add a new row to the table for each tenant and fill it with tenant data
    Dim d As Dictionary
    For Each d In GetTenantResponse.Data
        Set NewRow = TenantTable.ListRows.Add(Position:=1, AlwaysInsert:=True)
        With NewRow.Range
            .Cells(1, 1).value = Now
            ' DEBUG: Must call "Parent" first to avoid unexpected relative reference
            .Parent.Range(.Cells(1, 2), .Cells(1, d.Count + 1)) = d.Items
        End With
    Next d
                         
    ' Log success of storing tenant information
    WebHelpers.LogDebug "Stored current tenants info success, sheet: " & auth_SheetTenantsStorage, "XeroAuthenticator.StoreTenants"
    
    ' Limit the number of rows in the TenantTable to 50, deleting any excess rows from the bottom
    LimitTableRows TenantTable, 50
    
StoreTenants_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    Set NewRow = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while storing tenants info into Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.StoreTenants", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.StoreTenants", auth_ErrorDescription
    End If
End Sub

''
' Loads saved tenant information from an Excel sheet.
'
' @method LoadTenants
'
' This function performs the following steps:
' 1. Clears the existing dictionary of tenant IDs.
' 2. Checks if the worksheet containing tenant information exists. If not, logs a failure and exits the subroutine.
' 3. Checks if the ListObject (table) containing tenant information exists. If not, logs a failure and exits the subroutine.
' 4. Verifies if there are any rows in the table. If not, logs a failure and exits the subroutine.
' 5. Retrieves the most recent get date from the top row of the table for comparison.
' 6. Iterates through each row in the table, loading only the most recent tenant ID into the dictionary.
' 7. Logs the success of loading the tenants into the dictionary.
' 8. Handles any errors that occur during the process, logging the error details and raising an error.
''
Private Sub LoadTenants()
    On Error GoTo LoadTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewestGetDate As Date
    
    ' Clear the 'DTenantsId' property
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    Set wb = ThisWorkbook
    ' Check if the sheet for loading tenant informations exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Load tenants failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) for loading tenant informations exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Check if there are rows in the ListObject (table)
    If Not TenantTable.ListRows.Count > 0 Then
        ' Log failure if there are no rows in the table
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table) " & auth_TableTenantsStorage & " does not have any rows/records", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' The most recent record is at the top row (index = 1)
        ' Load the most recent get date for comparison later
        NewestGetDate = TenantTable.ListRows(1).Range.Cells(1, 1).value
    End If
    
    ' Load tenants into the 'DTenantsId' property
    Dim TenantRow As ListRow
    For Each TenantRow In TenantTable.ListRows
        ' Only load the most recent tenant ID
        If CDate(TenantRow.Range.Cells(1, 1).value) <> NewestGetDate Then
            Exit For
        End If
        
        ' Example row/record structure from TenantTable:
        ' Column Header: Value
        '  1    "Get_Datetime": "07/05/2024  17:06:11"
        '  2    "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1"
        '  3    "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d"
        '  4    "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056"
        '  5    "tenantType": "ORGANISATION"
        '  6    "tenantName": "Maple Florist"
        '  7    "createdDateUtc": "2019-07-09T23:40:30.1833130"
        '  8    "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        
        ' Based on StoreTenants method (10 May 2024)
        
        ' Store in dictionary with "tenantName" as the key and "tenantId" as the value
        DTenantsId(CStr(TenantRow.Range.Cells(1, 6).value)) = CStr(TenantRow.Range.Cells(1, 4).value)
    Next TenantRow
    
    ' Log success
    WebHelpers.LogDebug "LoadTenants success, available tenants: " & CStr(DTenantsId.Count), "XeroAuthenticator.LoadTenants"
        
LoadTenants_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    Set TenantRow = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while loading tenants from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.LoadTenants", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.LoadTenants", auth_ErrorDescription
    End If
End Sub

''
' Clears/deletes all saved tenant information from this Excel workbook.
'
' @method ClearTenants
'
' This function performs the following steps:
' 1. Checks if the sheet and table from which to clear tenant information exist.
' 2. If they do not exist, notifies the user and exits.
' 3. If found, proceeds to delete the table and its contents.
' 4. Notifies the user of the result, performs cleanup, and handles errors.
'
' Note: This function uses the 'auth_SheetTenantsStorage' and 'auth_TableTenantsStorage' constants.
''
Private Sub ClearTenants()
    On Error GoTo ClearTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    
    Set wb = ThisWorkbook
    ' Check if the sheet from which to clear tenant information exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Clear tenants cache failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify the user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the sheet does not exist
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) from which to clear tenant information exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Clear tenants cache failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify the user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the table does not exist
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Delete the table (ListObject) and all of its rows/records
    TenantTable.Delete
    
    ' Log success of clearing tenant information
    WebHelpers.LogDebug "Clearing tenants cache succeeded; all records of tenant information were deleted.", "XeroAuthenticator.ClearTenants"
    ' Notify the user of successful deletion
    MsgBox "Clearing Xero organizations cache succeeded; all records of Xero organizations were deleted.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ClearTenants_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing tenants cache from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearTenants", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearTenants", auth_ErrorDescription
    End If
End Sub

''
' Clears/deletes all saved tokens from this Excel workbook.
'
' @method ClearTokens
'
' This function performs the following steps:
' 1. Checks if the sheet from which to clear token information exists.
' 2. If the sheet does not exist, notifies the user and exits.
' 3. If the sheet is found, proceeds to check if the table (ListObject) exists.
' 4. If the table does not exist, notifies the user and exits.
' 5. If the table is found, deletes the table and its contents.
' 6. Notifies the user of the result, performs cleanup, and handles errors.
'
' Note: This function uses the 'auth_SheetTokensStorage' and 'auth_TableTokensStorage' constants.
''
Private Sub ClearTokens()
    On Error GoTo ClearTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    
    Set wb = ThisWorkbook
    ' Check if the sheet from which to clear token information exists
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Clear tokens cache failed, sheet: " & auth_SheetTokensStorage & "does not exist", "XeroAuthenticator.ClearTokens"
        ' Notify the user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the sheet does not exist
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check if the ListObject (table) from which to clear token information exists
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Clear tokens cache failed, ListObject(Table): " & auth_TableTokensStorage & " does not exist", "XeroAuthenticator.ClearTokens"
        ' Notify the user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the table does not exist
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    ' Delete the table (ListObject) and all of its rows/records
    TokenTable.Delete
    
    ' Log success of clearing token information
    WebHelpers.LogDebug "Clearing tokens cache succeeded; all records of tokens were deleted.", "XeroAuthenticator.ClearTokens"
    ' Notify the user of successful deletion
    MsgBox "Clearing tokens cache succeeded; all records of tokens were deleted.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ClearTokens_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing tokens cache from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearTokens", auth_ErrorDescription
    End If
End Sub

Private Sub Class_Initialize()
    Me.Scopes = Array("")
    LoadTokens
End Sub

