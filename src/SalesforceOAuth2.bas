Attribute VB_Name = "SalesforceOAuth2"
Option Explicit

' ==============================================================================
' SalesforceOAuth2.bas - OAuth 2.0 PKCE Orchestrator for VBA
' Implements the client flow including Code Verifier/Challenge handling,
' authorization URL assembly, token exchange, and refresh tokens.
' ==============================================================================

Private Const SUCCESS_URL_PREFIX As String = "https://login.salesforce.com/services/oauth2/success"

' ==============================================================================
' URL Encoder for Query Parameters
' ==============================================================================
Public Function URLEncode(ByVal StringVal As String) As String
    Dim i As Long
    Dim charCode As Integer
    Dim char As String
    Dim result As String
    
    result = ""
    For i = 1 To Len(StringVal)
        char = Mid$(StringVal, i, 1)
        charCode = Asc(char)
        
        Select Case charCode
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                ' Unreserved characters: A-Z, a-z, 0-9, '-', '.', '_', '~'
                result = result & char
            Case 32
                result = result & "%20"
            Case Else
                Dim hexStr As String
                hexStr = Hex(charCode)
                If Len(hexStr) = 1 Then hexStr = "0" & hexStr
                result = result & "%" & hexStr
        End Select
    Next i
    
    URLEncode = result
End Function

' ==============================================================================
' Assemble Salesforce Authorization URL
' ==============================================================================
Public Function GetAuthorizationUrl( _
    ByVal ClientId As String, _
    ByVal RedirectUri As String, _
    ByVal CodeChallenge As String, _
    Optional ByVal State As String = "", _
    Optional ByVal Domain As String = "login.salesforce.com" _
) As String

    Dim url As String
    url = "https://" & Domain & "/services/oauth2/authorize" & _
          "?response_type=code" & _
          "&client_id=" & URLEncode(ClientId) & _
          "&redirect_uri=" & URLEncode(RedirectUri) & _
          "&code_challenge=" & URLEncode(CodeChallenge) & _
          "&code_challenge_method=S256"
          
    If Len(State) > 0 Then
        url = url & "&state=" & URLEncode(State)
    End If
    
    GetAuthorizationUrl = url
End Function

' ==============================================================================
' Perform Authentication via Manual/Browser Copy-Paste Redirect
' Opens the default browser and prompts user to paste the callback URL/Code.
' ==============================================================================
Public Function AuthenticateManual( _
    ByVal ClientId As String, _
    ByVal RedirectUri As String, _
    ByRef OutCodeVerifier As String, _
    Optional ByVal Domain As String = "login.salesforce.com" _
) As String

    ' 1. Generate Verifier & Challenge
    Dim verifier As String
    verifier = CryptoUtils.GenerateCodeVerifier(64)
    OutCodeVerifier = verifier
    
    Dim challenge As String
    challenge = CryptoUtils.GenerateCodeChallenge(verifier)
    
    ' 2. Get auth URL
    Dim authUrl As String
    authUrl = GetAuthorizationUrl(ClientId, RedirectUri, challenge, "state123", Domain)
    
    ' 3. Open in default browser
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run authUrl
    
    ' 4. Prompt user to copy redirect URL back to Excel
    Dim promptMsg As String
    promptMsg = "A browser window has opened for you to log in to Salesforce." & vbCrLf & vbCrLf & _
                "1. Log in and authorize the application." & vbCrLf & _
                "2. You will be redirected to a callback page." & vbCrLf & _
                "3. Copy the ENTIRE URL from the browser's address bar and paste it below:"
                
    Dim responseUrl As String
    responseUrl = InputBox(promptMsg, "Paste Redirect URL", "")
    
    If Len(Trim(responseUrl)) = 0 Then
        Err.Raise vbObjectError + 3002, "SalesforceOAuth2", "Authentication was cancelled by the user."
    End If
    
    ' 5. Extract authorization code from the pasted URL
    Dim authCode As String
    authCode = ExtractCodeFromUrl(responseUrl)
    
    If Len(authCode) = 0 Then
        Err.Raise vbObjectError + 3003, "SalesforceOAuth2", "Failed to extract authorization code from URL. Make sure it contains 'code=' parameter."
    End If
    
    AuthenticateManual = authCode
End Function

' ==============================================================================
' Extract Code Parameter from URL
' ==============================================================================
Private Function ExtractCodeFromUrl(ByVal Url As String) As String
    Dim codeKey As String
    codeKey = "code="
    
    Dim startPos As Long
    startPos = InStr(1, Url, codeKey, vbTextCompare)
    
    If startPos = 0 Then
        ' Maybe the user pasted just the code itself
        If Len(Url) > 20 And InStr(1, Url, "/") = 0 And InStr(1, Url, "=") = 0 Then
            ExtractCodeFromUrl = Trim(Url)
            Exit Function
        End If
        ExtractCodeFromUrl = ""
        Exit Function
    End If
    
    startPos = startPos + Len(codeKey)
    
    Dim endPos As Long
    endPos = InStr(startPos, Url, "&")
    
    If endPos = 0 Then
        ExtractCodeFromUrl = Mid$(Url, startPos)
    Else
        ExtractCodeFromUrl = Mid$(Url, startPos, endPos - startPos)
    End If
    
    ' URL decode the code if necessary (replace %3D, %2F, etc.)
    ExtractCodeFromUrl = URLDecode(ExtractCodeFromUrl)
End Function

' ==============================================================================
' URL Decoder for Code extraction
' ==============================================================================
Private Function URLDecode(ByVal EncVal As String) As String
    Dim result As String
    result = EncVal
    result = Replace(result, "%2F", "/")
    result = Replace(result, "%3D", "=")
    result = Replace(result, "%2B", "+")
    result = Replace(result, "%25", "%")
    URLDecode = result
End Function

' ==============================================================================
' Post Request helper
' ==============================================================================
Private Function PostRequest(ByVal Url As String, ByVal Payload As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    http.Open "POST", Url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send Payload
    
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 3004, "SalesforceOAuth2", "OAuth request failed. HTTP " & http.Status & ": " & http.responseText
    End If
    
    PostRequest = http.responseText
End Function

' ==============================================================================
' Exchange Authorization Code for Access & Refresh Tokens
' ==============================================================================
Public Function RequestTokens( _
    ByVal ClientId As String, _
    ByVal RedirectUri As String, _
    ByVal CodeVerifier As String, _
    ByVal AuthCode As String, _
    Optional ByVal Domain As String = "login.salesforce.com" _
) As Object

    Dim tokenUrl As String
    tokenUrl = "https://" & Domain & "/services/oauth2/token"
    
    Dim payload As String
    payload = "grant_type=authorization_code" & _
              "&client_id=" & URLEncode(ClientId) & _
              "&redirect_uri=" & URLEncode(RedirectUri) & _
              "&code_verifier=" & URLEncode(CodeVerifier) & _
              "&code=" & URLEncode(AuthCode)
              
    Dim jsonResponse As String
    jsonResponse = PostRequest(tokenUrl, payload)
    
    Set RequestTokens = JSONParser.ParseJSON(jsonResponse)
End Function

' ==============================================================================
' Refresh Token to retrieve a new Access Token
' ==============================================================================
Public Function RefreshAccessToken( _
    ByVal ClientId As String, _
    ByVal RefreshToken As String, _
    Optional ByVal Domain As String = "login.salesforce.com" _
) As Object

    Dim tokenUrl As String
    tokenUrl = "https://" & Domain & "/services/oauth2/token"
    
    Dim payload As String
    payload = "grant_type=refresh_token" & _
              "&client_id=" & URLEncode(ClientId) & _
              "&refresh_token=" & URLEncode(RefreshToken)
              
    Dim jsonResponse As String
    jsonResponse = PostRequest(tokenUrl, payload)
    
    Set RefreshAccessToken = JSONParser.ParseJSON(jsonResponse)
End Function
