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

Public Property Get DTenantsId() As Dictionary
    If pDTenantsId Is Nothing Then
        Set pDTenantsId = New Dictionary
        pDTenantsId.CompareMode = TextCompare
    End If
    
    Set DTenantsId = pDTenantsId
End Property


' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ClientId
' @param {String} ClientSecret
''
Public Sub Setup(ClientId As String, ClientSecret As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
End Sub

''
' Login to Google
''
Public Sub Login()
    On Error GoTo auth_Cleanup
    
    Dim objBrowser As New CDPBrowser
    
    objBrowser.start name:="edge", appUrl:=Me.GetLoginUrl, cleanActive:=True, reAttach:=True
    
    ' Wait for login to complete
    Dim auth_IsApprove As Boolean
    Dim auth_IsError As Boolean
    
    Do
        auth_IsApprove = auth_LoginIsApproval(objBrowser)
        auth_IsError = auth_LoginIsError(objBrowser)
        objBrowser.sleep 1
    Loop Until auth_IsApprove Or auth_IsError

    If auth_IsError Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", auth_LoginExtractError(objBrowser)
    End If
    
    ' Success if extracted code is not empty, else the login might be cancelled by user
    Me.AuthorizationCode = auth_LoginExtractCode(objBrowser)
    If Me.AuthorizationCode = "" Then
        MsgBox "Login might be canceled! process is stopped", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was canceled"
    End If
    
auth_Cleanup:
    
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
' Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.AccessToken = ""
    Me.AccessTokenExpiration = CDate(0)
    Me.RefreshToken = ""
    Me.RefreshTokenExpiration = CDate(0)
End Sub

''
' Add scope to authorized scopes
'
' - To get a list of available scopes, visit https://developers.google.com/oauthplayground/
' - As a shortcut, if a domain isn't given, https://www.googleapis.com/auth/ is automatically added
'
' @example
' Auth.AddScope "yt-analytics.readonly" ' -> https://www.googleapis.com/auth/yt-analytics.readonly
' Auth.AddScope "https://mail.google.com/"
'
' @param {String} Scope
''
Public Sub AddScope(Scope As String)
    Dim auth_Scopes As Variant
    
    ' Prepare scopes array
    auth_Scopes = Me.Scopes
    If VBA.IsEmpty(auth_Scopes) Then
        ReDim auth_Scopes(0 To 0)
    Else
        ReDim Preserve auth_Scopes(0 To UBound(auth_Scopes) + 1)
    End If
    
    ' Update "Scopes" array property
    auth_Scopes(UBound(auth_Scopes)) = Scope
    Me.Scopes = auth_Scopes
End Sub

''
' Get token (either new or refreshed)
'
' @param {WebClient} Client
''
Public Sub GetToken(Client As WebClient)
        
    ' No need to process GetToken if access token exist and not expired
    If Me.AccessToken <> "" And Not Me.AccessTokenExpiration < Now Then
        WebHelpers.LogDebug "AccessToken already exist and not expired, skipping GetToken process", "XeroAuthenticator.GetToken"
        Exit Sub
    ' Get refreshed token if refresh token exist and not expired
    ElseIf Me.RefreshToken <> "" And Not Me.RefreshTokenExpiration < Now Then
        WebHelpers.LogDebug "RefreshToken exist and not expired, attempting to refresh AccessToken", "XeroAuthenticator.GetToken"
        GetRefreshedToken Client
    ' Get new token
    Else
        ' Login first if auth code doesn't exist
        If Me.AuthorizationCode = "" Then
            WebHelpers.printMsg "XeroAuthenticator.GetToken", "Auth code not found, redirecting to Login"
            Me.Login
        End If
        WebHelpers.LogDebug "Tokens not found or expired, getting new tokens", "XeroAuthenticator.GetToken"
        GetNewToken Client
    End If
    ' Store new or refreshed tokens
    StoreTokens
    
End Sub

''
' Get tenant - request to Xero API
'
' @param {WebClient} Client
''
Public Sub GetTenant(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TenantClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clear dictionary
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TenantClient = Client.Clone
    Set auth_TenantClient.Authenticator = Nothing
    auth_TenantClient.BaseUrl = "https://api.xero.com/"
    
    ' Prepare token request
    auth_Request.Resource = "connections"
    auth_Request.Method = WebMethod.HttpGet
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    Set auth_Response = auth_TenantClient.Execute(auth_Request)
    
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
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
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

Public Function SelectTenant(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    ' Load list of tenants from saved excel sheet
    LoadTenants
    
    ' If load results doesn't exits, try get request
    If Not DTenantsId.Count > 0 Then
        GetTenant Client
    End If

    ' Initialize form
    Dim SelectForm1 As SelectXeroOrgForm
    Set SelectForm1 = New SelectXeroOrgForm
    
    ' Assign tenants list to form
    SelectForm1.OrgList = DTenantsId.Keys
    
    ' Show form to user
    SelectForm1.show
    
    ' If form is hide or closed then it will resume
    ' Check if there is an update request
    If SelectForm1.UserRequestUpdate Then
        ' Get and update the tenant list
        GetTenant Client
        SelectForm1.OrgList = DTenantsId.Keys
        
        ' Reset the form var to false
        SelectForm1.UserRequestUpdate = False
        
        ' Show the form again with the updated list
        SelectForm1.show
    End If
    
    ' Check if user canceled the form
    If SelectForm1.UserCancel Then
        ' Notify user and raise error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Err.Raise 11040 + vbObjectError, "SelectXeroOrgForm", "User canceled selection form"
    End If
    
    SelectTenant = DTenantsId(CStr(SelectForm1.ComboBox1.value))

auth_Cleanup:
    ' Unload when everything is finished
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
' Clears/delete all saved cache (tenants & tokens) from this Excel workbook
'
''
Public Sub ClearAllCache(isClearTenant As Boolean, isClearToken As Boolean)
    On Error GoTo auth_Cleanup
    
    If isClearTenant Then
        ClearTenants
    End If
    
    If isClearToken Then
        ClearTokens
    End If
    
auth_Cleanup:
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearAllCache", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearAllCache", auth_ErrorDescription
    End If
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' GetToken process will check current tokens exist and not expired, else will get new tokens
    Me.GetToken Client
    
    ' Set up required headers for API calls - AccessToken
    Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    ' Select tenant id to be included in the request
    Dim SelectedTenantId As String
    SelectedTenantId = SelectTenant(Client)
    
    ' Set up required headers for API calls - Xero-tenant-id
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
' Get login url for current scopes
'
' @internal
' @return {String}
''
Public Function GetLoginUrl() As String
    ' Use Request for Url helpers
    Dim auth_Request As New WebRequest
    auth_Request.Resource = auth_AuthorizationUrl
    
    auth_Request.AddQuerystringParam "response_type", "code"
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "state", "123"
    
    If Not VBA.IsEmpty(Me.Scopes) Then
        auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, " ")
    Else
        auth_Request.AddQuerystringParam "scope", ""
    End If
    
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_LoginIsApproval(objBrowser As CDPBrowser) As Boolean
    Dim auth_BrowserUrlParts As Dictionary
    Dim auth_RedirectUrlParts As Dictionary
    
    Set auth_BrowserUrlParts = WebHelpers.GetUrlParts(objBrowser.url)
    Set auth_RedirectUrlParts = WebHelpers.GetUrlParts(auth_RedirectUrl)
    
    If auth_BrowserUrlParts("Host") = auth_RedirectUrlParts("Host") Then auth_LoginIsApproval = True
End Function

Private Function auth_LoginIsError(objBrowser As CDPBrowser) As Boolean
    Dim auth_Element As CDPElement
    
    objBrowser.wait till:=ReadyState.isComplete
    ' Search 'error' word in the html title
    Set auth_Element = objBrowser.getElementByQuery("title", True)
    If auth_Element.isExist Then
        auth_LoginIsError = InStr(1, CStr(auth_Element.innerText), "error", vbTextCompare) > 0
    End If
End Function

Private Function auth_LoginExtractCode(objBrowser As CDPBrowser) As String
    If auth_LoginIsApproval(objBrowser) Then
        ' Extract authorization code
        Dim auth_UrlParts As Dictionary
        Dim auth_ParsedUrlParams As Dictionary
        
        Set auth_UrlParts = WebHelpers.GetUrlParts(objBrowser.url)
        Set auth_ParsedUrlParams = ParseUrlEncoded(auth_UrlParts("Querystring"))
        auth_LoginExtractCode = auth_ParsedUrlParams("code")
    End If
End Function

Private Function auth_LoginExtractError(objBrowser As CDPBrowser) As String
    Dim auth_Element As CDPElement
    ' Extract error code e.g. "Error code: 500"
    Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]")
    If auth_Element.isExist Then
        auth_LoginExtractError = auth_Element.innerText
        ' Extract error description
        Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]/following-sibling::p")
        If auth_Element.isExist Then
            auth_LoginExtractError = auth_LoginExtractError & " - " & auth_Element.innerText
        End If
    Else
        auth_LoginExtractError = "Unknown"
    End If
End Function

''
' Get new token (for current AuthorizationCode)
'
' @param {WebClient} Client
''
Private Sub GetNewToken(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://identity.xero.com/"
    
    ' Prepare token request
    auth_Request.Resource = "connect/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "authorization", "Basic " & Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare request body
    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "redirect_uri", auth_RedirectUrl
    auth_Body.Add "grant_type", "authorization_code"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Response 401 - auth code might be expired
    ' Attemp to re-login and update auth code, then repeat request once if permitted
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
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetNewToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
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
' Get refreshed token - in case of expired access token
'
' @param {WebClient} Client
''
Private Sub GetRefreshedToken(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://identity.xero.com/"
    
    ' Prepare token request
    auth_Request.Resource = "connect/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "authorization", "Basic " + Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare request body
    auth_Body.Add "refresh_token", Me.RefreshToken
    auth_Body.Add "grant_type", "refresh_token"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        Me.AccessToken = auth_Response.Data("access_token")
        Me.AccessTokenExpiration = DateAdd("s", CLng(auth_Response.Data("expires_in")), Now)
        Me.RefreshToken = auth_Response.Data("refresh_token")
        Me.RefreshTokenExpiration = DateAdd("d", 60, Now)
        
        WebHelpers.LogDebug "New Access Token: " & Me.AccessToken & vbNewLine & vbTab & _
            "New Refresh Token: " & Me.RefreshToken, "XeroAuthenticator.GetNewToken"
    
    ' Response 401 - refresh token might be expired
    ' Attemp to re-login and update auth code if permitted
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
                
                ' Clean up and redirect the process to get new token procedure
                Set auth_TokenClient = Nothing
                Set auth_Request = Nothing
                Set auth_Response = Nothing
                
                GetNewToken Client
            Case vbNo
                ' Notify stopping refresh token process
                WebHelpers.printMsg "XeroAuthenticator.GetRefreshedToken", "Abort refreshing token process"
        End Select
    Else
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetRefreshedToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
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
' Save obtained access token and refresh token into an Excel sheet
'
''
Private Sub StoreTokens()
    On Error GoTo StoreTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    Dim NewRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Create new sheet to store tokens if not exist already
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        Set sh = wb.Sheets.Add(Before:=wb.Sheets(1))
        sh.name = auth_SheetTokensStorage
        
        'Hide sheet from user
        sh.Visible = xlSheetVeryHidden
    Else
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check whether ListObject(Table) to store tokens exist
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        'Create the table headers for storing tokens
        sh.Cells(1, 1).value = "Get_Datetime"
        sh.Cells(1, 2).value = "Access_Token"
        sh.Cells(1, 3).value = "AT_Expired_Datetime"
        sh.Cells(1, 4).value = "Refresh_Token"
        sh.Cells(1, 5).value = "RT_Expired_Datetime"
        
        'Setup the new table as excel table(ListObject)
        Set TokenTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                        source:=sh.Range("A1:E1"), _
                                        XlListObjectHasHeaders:=xlYes)
        TokenTable.name = auth_TableTokensStorage
    Else
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    'Add new row and fill
    Set NewRow = TokenTable.ListRows.Add(Position:=1, AlwaysInsert:=True)
    NewRow.Range = Array(Now, _
                         Me.AccessToken, _
                         Me.AccessTokenExpiration, _
                         Me.RefreshToken, _
                         Me.RefreshTokenExpiration)
                         
    'Log success
    WebHelpers.LogDebug "Stored existing tokens success, sheet: " & auth_SheetTokensStorage, "XeroAuthenticator.StoreTokens"
    
    'Limit TokenTable rows/records to 20 rows only, delete the remains starting from bottom
    LimitTableRows TokenTable, 20
    
StoreTokens_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    Set NewRow = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while storing tokens into Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.StoreTokens", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.StoreTokens", auth_ErrorDescription
    End If
End Sub

''
' Load saved access token and refresh token from an Excel sheet
'
''
Private Sub LoadTokens()
    On Error GoTo LoadTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    Dim NewestRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Check whether sheet to load tokens from exist
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        ' Log failure due to sheet doesn't exist
        WebHelpers.LogDebug "Load tokens failed, sheet: " & auth_SheetTokensStorage & "does not exist", "XeroAuthenticator.LoadTokens"
        ' Exit sub
        Exit Sub
    Else
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check whether ListObject(Table) to load tokens from exist
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        ' Log failure due to table doesn't exist
        WebHelpers.LogDebug "Load tokens failed, ListObject(Table): " & auth_TableTokensStorage & " does not exist", "XeroAuthenticator.LoadTokens"
        ' Exit sub
        Exit Sub
    Else
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    ' Check whether row(s) in ListObject(Table) to load tokens from exist
    If Not TokenTable.ListRows.Count > 0 Then
        ' Log failure due to row(s) in the table doesn't exist
        WebHelpers.LogDebug "Load tokens failed, ListObject(Table) " & auth_TableTokensStorage & " does not have any rows/records", "XeroAuthenticator.LoadTokens"
        ' Exit sub
        Exit Sub
    Else
        ' Most recent record of tokens are placed at the top row (index = 1)
        Set NewestRow = TokenTable.ListRows(1)
    End If
    
    ' Load tokens into auth object
    Me.AccessToken = CStr(NewestRow.Range.Cells(1, 2).value)
    Me.AccessTokenExpiration = CDate(NewestRow.Range.Cells(1, 3).value)
    Me.RefreshToken = CStr(NewestRow.Range.Cells(1, 4).value)
    Me.RefreshTokenExpiration = CDate(NewestRow.Range.Cells(1, 5).value)
    
    ' Post-process, check expired tokens (access token)
    If Me.AccessTokenExpiration < Now Then
        ' Log failure due to token expired
        WebHelpers.LogDebug "Load AccessToken failed, AccessToken expired at: " & Format(Me.AccessTokenExpiration, "dd/mm/yyyy hh:mm"), "XeroAuthenticator.LoadTokens"
        ' Revert object values to empty or default
        Me.AccessToken = ""
        Me.AccessTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load AccessToken success: " & Me.AccessToken, "XeroAuthenticator.LoadTokens"
    End If
    
    ' Check expired tokens (refresh token)
    If Me.RefreshTokenExpiration < Now Then
        ' Log failure due to token expired
        WebHelpers.LogDebug "Load RefreshToken failed, RefreshToken expired at: " & Format(Me.RefreshTokenExpiration, "dd/mm/yyyy hh:mm"), "XeroAuthenticator.LoadTokens"
        ' Revert object values to empty or default
        Me.RefreshToken = ""
        Me.RefreshTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load RefreshToken success: " & Me.RefreshToken, "XeroAuthenticator.LoadTokens"
    End If
        
LoadTokens_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    Set NewestRow = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while loading tokens from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.LoadTokens", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.LoadTokens", auth_ErrorDescription
    End If
End Sub

''
' Limit rows in table, delete rows above certain counts
' Used for limiting tokens table rows/records - inside StoreTokens procedure
'
' @param {ListObject} Table to limit
' @param {Integer} Rows limit number
''
Private Sub LimitTableRows(Table As ListObject, IRowsLimit As Integer)
    
    Dim TblRows As ListRows
    Set TblRows = Table.ListRows
    
    ' Exit sub if limit or table rows is 0
    If Not IRowsLimit > 0 Or Not TblRows.Count > 0 Then
        WebHelpers.LogDebug "Table " & Table.name & " - limit number or table rows can not be zero - " & _
                            "IRowsLimit: " & IRowsLimit & _
                            ", TblRows.Count: " & TblRows.Count, _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    ' Exit sub if the table rows is already under/equal to the limit
    ElseIf Not TblRows.Count > IRowsLimit Then
        WebHelpers.LogDebug "Table " & Table.name & " - rows is already under/equal to the limit", _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    End If
    
    ' Loop while table rows count is bigger than the limit number
    Do While TblRows.Count > IRowsLimit
        ' Delete from table bottom row
        TblRows(TblRows.Count).Delete
    Loop
    
    WebHelpers.LogDebug "Table " & Table.name & " - successfully limited to only " & TblRows.Count & " rows", _
                        "XeroAuthenticator.LimitTableRows"

End Sub

''
' Save obtained Xero tenants informations into an Excel sheet
'
' @param {WebResponse} Response received from GetTenant procedure
''
Private Sub StoreTenants(GetTenantResponse As WebResponse)
    On Error GoTo StoreTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Create new sheet to store tenants info if not exist already
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        Set sh = wb.Sheets.Add(Before:=wb.Sheets(1))
        sh.name = auth_SheetTenantsStorage
        
        'Hide sheet from user
        sh.Visible = xlSheetVeryHidden
    Else
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check whether ListObject(Table) to store tenants info exist
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        'Create the table headers for storing tenants info
        'Expected 'Data(1)' example from get tenants response:
        '   "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1",
        '   "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d",
        '   "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056",
        '   "tenantType": "ORGANISATION",
        '   "tenantName": "Maple Florist",
        '   "createdDateUtc": "2019-07-09T23:40:30.1833130",
        '   "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        '
        'Based on official docs (10 May 2024): https://developer.xero.com/documentation/guides/oauth2/auth-flow/#5-check-the-tenants-youre-authorized-to-access
        'Column to record GetTenant response datetime
        sh.Cells(1, 1).value = "Get_Datetime"
        'Columns based on response Data Keys
        sh.Range(sh.Cells(1, 2), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)) = GetTenantResponse.Data(1).Keys
        
        'Setup the new table as excel table(ListObject)
        Set TenantTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                            source:=sh.Range(sh.Cells(1, 1), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)), _
                                            XlListObjectHasHeaders:=xlYes)
        TenantTable.name = auth_TableTenantsStorage
    Else
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    'Add new row and fill
    Dim d As Dictionary
    For Each d In GetTenantResponse.Data
        Set NewRow = TenantTable.ListRows.Add(Position:=1, AlwaysInsert:=True)
        With NewRow.Range
            .Cells(1, 1).value = Now
            ' DEBUG: Must call "Parent" first to avoid unexpected relative reference
            .Parent.Range(.Cells(1, 2), .Cells(1, d.Count + 1)) = d.Items
        End With
    Next d
                         
    'Log success
    WebHelpers.LogDebug "Stored current tenants info success, sheet: " & auth_SheetTenantsStorage, "XeroAuthenticator.StoreTenants"
    
    'Limit TenantTable rows/records to 50 rows only, delete the remains starting from bottom
    LimitTableRows TenantTable, 50
    
StoreTenants_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    Set NewRow = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while storing tenants info into Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.StoreTenants", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.StoreTenants", auth_ErrorDescription
    End If
End Sub

''
' Load saved tenants info from an Excel sheet
'
''
Private Sub LoadTenants()
    On Error GoTo LoadTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewestGetDate As Date
    
    ' Clear dictionary
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    Set wb = ThisWorkbook
    ' Check whether sheet to load tenants info from exist
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to sheet doesn't exist
        WebHelpers.LogDebug "Load tenants failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit sub
        Exit Sub
    Else
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check whether ListObject(Table) to load tenants info from exist
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to table doesn't exist
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit sub
        Exit Sub
    Else
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Check whether row(s) in ListObject(Table) to load tenants info from exist
    If Not TenantTable.ListRows.Count > 0 Then
        ' Log failure due to row(s) in the table doesn't exist
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table) " & auth_TableTenantsStorage & " does not have any rows/records", "XeroAuthenticator.LoadTenants"
        ' Exit sub
        Exit Sub
    Else
        ' Most recent record of tenants info are placed at the top row (index = 1)
        ' Load most recent get date value for comparison later
        NewestGetDate = TenantTable.ListRows(1).Range.Cells(1, 1).value
    End If
    
    ' Load tenants into auth object
    Dim TenantRow As ListRow
    For Each TenantRow In TenantTable.ListRows
        ' Only load most recent tenants id
        If CDate(TenantRow.Range.Cells(1, 1).value) <> NewestGetDate Then
            Exit For
        End If
        
        ' Expected row/record example from TenantTable:
        ' Column Header: Value
        '  1    "Get_Datetime": "07/05/2024  17:06:11"
        '  2    "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1",
        '  3    "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d",
        '  4    "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056",
        '  5    "tenantType": "ORGANISATION",
        '  6    "tenantName": "Maple Florist",
        '  7    "createdDateUtc": "2019-07-09T23:40:30.1833130",
        '  8    "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        '
        ' Based on StoreTenants method (10 May 2024)
        
        ' Load into dictionary with "tenantName" as key, and "tenantId" as item/value
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
' Clears/delete all saved tenants info from this Excel workbook
'
''
Private Sub ClearTenants()
    On Error GoTo ClearTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    
    Set wb = ThisWorkbook
    ' Check whether sheet to clear tenants info from exist
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to sheet doesn't exist
        WebHelpers.LogDebug "Clear tenants cache failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit sub
        Exit Sub
    Else
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check whether ListObject(Table) to clear tenants info from exist
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to table doesn't exist
        WebHelpers.LogDebug "Clear tenants cache failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit sub
        Exit Sub
    Else
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Delete table (ListObject) and all of its rows/records
    TenantTable.Delete
    
    ' Log success
    WebHelpers.LogDebug "Clearing tenants cache succeeded; all records of tenant information were deleted.", "XeroAuthenticator.ClearTenants"
    ' Notify user
    MsgBox "Clearing Xero organizations cache succeeded; all records of Xero organizations were deleted.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ClearTenants_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing tenants cache from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearTenants", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearTenants", auth_ErrorDescription
    End If
End Sub

''
' Clears/delete all saved tokens from this Excel workbook
'
''
Private Sub ClearTokens()
    On Error GoTo ClearTokens_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TokenTable As ListObject
    
    Set wb = ThisWorkbook
    ' Check whether sheet to clear tokens info from exist
    If Not WebHelpers.WorksheetExists(auth_SheetTokensStorage, wb) Then
        ' Log failure due to sheet doesn't exist
        WebHelpers.LogDebug "Clear tokens cache failed, sheet: " & auth_SheetTokensStorage & "does not exist", "XeroAuthenticator.ClearTokens"
        ' Notify user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit sub
        Exit Sub
    Else
        Set sh = wb.Sheets(auth_SheetTokensStorage)
    End If
    
    ' Check whether ListObject(Table) to clear tokens info from exist
    If Not WebHelpers.ListObjectExists(auth_TableTokensStorage, sh) Then
        ' Log failure due to table doesn't exist
        WebHelpers.LogDebug "Clear tokens cache failed, ListObject(Table): " & auth_TableTokensStorage & " does not exist", "XeroAuthenticator.ClearTokens"
        ' Notify user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit sub
        Exit Sub
    Else
        Set TokenTable = sh.ListObjects(auth_TableTokensStorage)
    End If
    
    ' Delete table (ListObject) and all of its rows/records
    TokenTable.Delete
    
    ' Log success
    WebHelpers.LogDebug "Clearing tokens cache succeeded; all records of tokens were deleted.", "XeroAuthenticator.ClearTokens"
    ' Notify user
    MsgBox "Clearing tokens cache succeeded; all records of tokens were deleted.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ClearTokens_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TokenTable = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing tokens cache from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearTokens", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearTokens", auth_ErrorDescription
    End If
End Sub

Private Sub Class_Initialize()
    Me.Scopes = Array("")
    LoadTokens
End Sub

