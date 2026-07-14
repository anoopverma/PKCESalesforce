VERSION 5.00
Begin {C62DE69F-2E85-11D1-B28A-00C04F93801D} frmOAuthLogin 
   Caption         =   "Log in to Salesforce"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOAuthLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' frmOAuthLogin.frm - Interactive WebBrowser UserForm (Optional)
' Intercepts Salesforce OAuth authentication redirect code
' ==============================================================================

Public AuthCode As String
Public RedirectUri As String
Public IsCancelled As Boolean

' Dynamically hold references
Private WithEvents Browser As SHDocVw.WebBrowser
Attribute Browser.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    IsCancelled = True
    AuthCode = ""
End Sub

Public Sub StartLogin(ByVal AuthUrl As String, ByVal RedirectUrl As String)
    RedirectUri = RedirectUrl
    
    ' Add WebBrowser control dynamically at runtime
    Dim browserCtrl As Object
    On Error Resume Next
    Set browserCtrl = Me.Controls.Add("Shell.Explorer.2", "WebBrowser1", True)
    On Error GoTo 0
    
    If browserCtrl Is Nothing Then
        MsgBox "Failed to load ActiveX WebBrowser control." & vbCrLf & _
               "Falling back to Manual/System Browser Login mode.", vbInformation, "Notice"
        Me.Hide
        Exit Sub
    End If
    
    ' Layout browser control to fill form
    browserCtrl.Left = 6
    browserCtrl.Top = 6
    browserCtrl.Width = Me.InsideWidth - 12
    browserCtrl.Height = Me.InsideHeight - 12
    
    ' Hook the browser events
    Set Browser = browserCtrl.Object
    
    ' Navigate to Salesforce auth URL
    Browser.Navigate AuthUrl
    Me.Show
End Sub

Private Sub Browser_BeforeNavigate2( _
    ByVal pDisp As Object, _
    ByRef URL As Variant, _
    ByRef Flags As Variant, _
    ByRef TargetFrameName As Variant, _
    ByRef PostData As Variant, _
    ByRef Headers As Variant, _
    ByRef Cancel As Boolean _
)
    Dim currentUrl As String
    currentUrl = CStr(URL)
    
    ' Check if URL matches redirect URI
    If InStr(1, currentUrl, RedirectUri, vbTextCompare) = 1 Then
        Cancel = True
        IsCancelled = False
        
        ' Extract the code parameter
        Dim codeKey As String
        codeKey = "code="
        Dim startPos As Long
        startPos = InStr(1, currentUrl, codeKey, vbTextCompare)
        
        If startPos > 0 Then
            startPos = startPos + Len(codeKey)
            Dim endPos As Long
            endPos = InStr(startPos, currentUrl, "&")
            If endPos = 0 Then
                AuthCode = Mid$(currentUrl, startPos)
            Else
                AuthCode = Mid$(currentUrl, startPos, endPos - startPos)
            End If
            
            ' Unescape URL encoding
            AuthCode = Replace(AuthCode, "%2F", "/")
            AuthCode = Replace(AuthCode, "%3D", "=")
            AuthCode = Replace(AuthCode, "%2B", "+")
            AuthCode = Replace(AuthCode, "%25", "%")
        End If
        
        Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' User closed the window using [X]
        IsCancelled = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    Dim browserCtrl As Object
    Set browserCtrl = Me.Controls("WebBrowser1")
    If Not browserCtrl Is Nothing Then
        browserCtrl.Width = Me.InsideWidth - 12
        browserCtrl.Height = Me.InsideHeight - 12
    End If
    On Error GoTo 0
End Sub
