Attribute VB_Name = "OAuthDemo"
Option Explicit

' ==============================================================================
' OAuthDemo.bas - Salesforce PKCE Integration Demo
' Demonstrates how to initialize credentials, perform manual authentication,
' exchange authorization codes, run queries, and refresh tokens.
' ==============================================================================

' Change these to match your Salesforce Connected App settings!
Private Const CLIENT_ID As String = "YOUR_CONSUMER_KEY_HERE"
Private Const REDIRECT_URI As String = "https://login.salesforce.com/services/oauth2/success"
Private Const LOGIN_DOMAIN As String = "login.salesforce.com" ' Use "test.salesforce.com" for Sandbox

' State variables (in production, store these securely in a hidden worksheet or database)
Private mAccessToken As String
Private mRefreshToken As String
Private mInstanceUrl As String

' ==============================================================================
' Macro to Run the Full Salesforce PKCE OAuth Login and Test Flow
' ==============================================================================
Public Sub RunSalesforceOAuthDemo()
    Dim verifier As String
    Dim authCode As String
    Dim tokens As Object
    
    ' 1. Check if CLIENT_ID is filled
    If CLIENT_ID = "YOUR_CONSUMER_KEY_HERE" Then
        MsgBox "Please open the OAuthDemo module in the VBA Editor and replace CLIENT_ID with your Salesforce Connected App Consumer Key.", vbExclamation, "Configuration Required"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' 2. Authenticate the User and Capture the Authorization Code
    MsgBox "A browser window will open to authenticate with Salesforce. " & vbCrLf & _
           "After log in, please copy the full URL of the final redirect page and paste it into the next dialog.", vbInformation, "Start Authentication"
           
    authCode = SalesforceOAuth2.AuthenticateManual(CLIENT_ID, REDIRECT_URI, verifier, LOGIN_DOMAIN)
    
    ' 3. Exchange Authorization Code for Tokens
    Debug.Print "Exchanging authorization code for tokens..."
    Set tokens = SalesforceOAuth2.RequestTokens(CLIENT_ID, REDIRECT_URI, verifier, authCode, LOGIN_DOMAIN)
    
    ' 4. Extract and Save Tokens
    mAccessToken = tokens("access_token")
    mRefreshToken = tokens("refresh_token")
    mInstanceUrl = tokens("instance_url")
    
    Debug.Print "Access Token: " & Left$(mAccessToken, 15) & "..."
    Debug.Print "Refresh Token: " & Left$(mRefreshToken, 15) & "..."
    Debug.Print "Instance URL: " & mInstanceUrl
    
    MsgBox "Authentication Successful!" & vbCrLf & _
           "Connected to Instance: " & mInstanceUrl & vbCrLf & _
           "Testing API query next...", vbInformation, "Success"
           
    ' 5. Run a Test REST API Query (Fetch user name and email)
    Dim queryResultJson As String
    Dim queryResult As Object
    Dim records As Object
    Dim firstRecord As Object
    
    queryResultJson = SalesforceAPI.Query(mInstanceUrl, mAccessToken, "SELECT Name, Email FROM User LIMIT 1")
    Set queryResult = JSONParser.ParseJSON(queryResultJson)
    Set records = queryResult("records")
    
    If records.Count > 0 Then
        Set firstRecord = records(1)
        MsgBox "Salesforce API Test Query Succeeded!" & vbCrLf & vbCrLf & _
               "Current User: " & firstRecord("Name") & vbCrLf & _
               "Email: " & firstRecord("Email"), vbInformation, "Query Completed"
    Else
        MsgBox "Query ran successfully but returned 0 records.", vbInformation, "No Records"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Authentication Failed"
End Sub

' ==============================================================================
' Macro to Test Token Refresh Flow
' ==============================================================================
Public Sub RefreshSalesforceTokenDemo()
    If Len(mRefreshToken) = 0 Then
        MsgBox "Please run the 'RunSalesforceOAuthDemo' macro first to obtain a refresh token.", vbExclamation, "No Refresh Token"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Debug.Print "Refreshing access token..."
    Dim tokens As Object
    Set tokens = SalesforceOAuth2.RefreshAccessToken(CLIENT_ID, mRefreshToken, LOGIN_DOMAIN)
    
    mAccessToken = tokens("access_token")
    Debug.Print "New Access Token: " & Left$(mAccessToken, 15) & "..."
    
    MsgBox "Access token refreshed successfully!", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Refresh Failed"
End Sub
