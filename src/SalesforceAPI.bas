Attribute VB_Name = "SalesforceAPI"
Option Explicit

' ==============================================================================
' SalesforceAPI.bas - Authenticated REST API Client for VBA
' Provides methods to send authenticated HTTP requests to the Salesforce REST API.
' ==============================================================================

' ==============================================================================
' Execute REST API Request to Salesforce
' ==============================================================================
Public Function ExecuteRequest( _
    ByVal Method As String, _
    ByVal InstanceUrl As String, _
    ByVal Endpoint As String, _
    ByVal AccessToken As String, _
    Optional ByVal Body As String = "" _
) As String

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Construct complete URL
    Dim fullUrl As String
    If Right$(InstanceUrl, 1) = "/" And Left$(Endpoint, 1) = "/" Then
        fullUrl = InstanceUrl & Mid$(Endpoint, 2)
    ElseIf Right$(InstanceUrl, 1) <> "/" And Left$(Endpoint, 1) <> "/" Then
        fullUrl = InstanceUrl & "/" & Endpoint
    Else
        fullUrl = InstanceUrl & Endpoint
    End If
    
    ' Open connection
    http.Open Method, fullUrl, False
    
    ' Set standard Salesforce OAuth2 / JSON headers
    http.setRequestHeader "Authorization", "Bearer " & AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    
    ' Send payload if present
    If Len(Body) > 0 And (Method = "POST" Or Method = "PATCH" Or Method = "PUT") Then
        http.Send Body
    Else
        http.Send
    End If
    
    ' Error handling for unauthorized or bad requests
    If http.Status = 401 Then
        Err.Raise vbObjectError + 4001, "SalesforceAPI", "Session expired or unauthorized (401). Access Token needs refresh."
    ElseIf http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 4002, "SalesforceAPI", "API Request failed. Status: " & http.Status & " - " & http.responseText
    End If
    
    ExecuteRequest = http.responseText
End Function

' ==============================================================================
' Run SOQL Query Helper
' ==============================================================================
Public Function Query( _
    ByVal InstanceUrl As String, _
    ByVal AccessToken As String, _
    ByVal Soql As String, _
    Optional ByVal ApiVersion As String = "v60.0" _
) As String

    Dim queryEndpoint As String
    queryEndpoint = "/services/data/" & ApiVersion & "/query?q=" & SalesforceOAuth2.URLEncode(Soql)
    
    Query = ExecuteRequest("GET", InstanceUrl, queryEndpoint, AccessToken)
End Function
