Attribute VB_Name = "VBA_Tests"
Option Explicit

' ==============================================================================
' VBA_Tests.bas - Unit Tests for Cryptography and JSON Parser
' Verify correctness of SHA256, Base64URL, and JSON Parser before using them.
' ==============================================================================

' Helper to convert Byte array to lowercase Hex string
Private Function BytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim result As String
    result = ""
    For i = 0 To UBound(Bytes)
        result = result & Right$("0" & Hex(Bytes(i)), 2)
    Next i
    BytesToHex = LCase$(result)
End Function

' ==============================================================================
' Run All Unit Tests
' ==============================================================================
Public Sub RunAllTests()
    Dim passed As Long
    Dim failed As Long
    
    Debug.Print "========================================="
    Debug.Print "RUNNING SALESFORCE PKCE OAUTH UNIT TESTS"
    Debug.Print "========================================="
    
    ' Test 1: SHA256 Hashing
    If TestSHA256() Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
    
    ' Test 2: Base64URL Encoding
    If TestBase64URL() Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
    
    ' Test 3: Code Verifier & Challenge
    If TestCodeVerifierAndChallenge() Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
    
    ' Test 4: JSON Parsing
    If TestJSONParser() Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
    
    Debug.Print "-----------------------------------------"
    Debug.Print "RESULTS: " & passed & " Passed, " & failed & " Failed."
    Debug.Print "========================================="
    
    If failed = 0 Then
        MsgBox "All unit tests completed successfully!", vbInformation, "Tests Passed"
    Else
        MsgBox "Some unit tests failed. Check the Immediate Window (Ctrl + G) for details.", vbCritical, "Tests Failed"
    End If
End Sub

' ==============================================================================
' Test SHA256 Hashing
' ==============================================================================
Private Function TestSHA256() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test Vector 1: SHA256("abc")
    Dim hash1() As Byte
    hash1 = CryptoUtils.GetSHA256Hash("abc")
    Dim hex1 As String
    hex1 = BytesToHex(hash1)
    
    Const EXPECTED_HEX_1 As String = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    
    If hex1 <> EXPECTED_HEX_1 Then
        Debug.Print "TestSHA256 Fail: Expected " & EXPECTED_HEX_1 & " but got " & hex1
        TestSHA256 = False
        Exit Function
    End If
    
    ' Test Vector 2: SHA256("")
    Dim hash2() As Byte
    hash2 = CryptoUtils.GetSHA256Hash("")
    Dim hex2 As String
    hex2 = BytesToHex(hash2)
    
    Const EXPECTED_HEX_2 As String = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    
    If hex2 <> EXPECTED_HEX_2 Then
        Debug.Print "TestSHA256 Fail: Expected " & EXPECTED_HEX_2 & " but got " & hex2
        TestSHA256 = False
        Exit Function
    End If
    
    Debug.Print "TestSHA256: PASSED"
    TestSHA256 = True
    Exit Function

ErrorHandler:
    Debug.Print "TestSHA256 Error: " & Err.Description
    TestSHA256 = False
End Function

' ==============================================================================
' Test Base64URL Encoding
' ==============================================================================
Private Function TestBase64URL() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test Base64URL encoding with padding removal and character replacement
    Dim testBytes(0 To 4) As Byte
    testBytes(0) = &HFC
    testBytes(1) = &H3F
    testBytes(2) = &H90
    testBytes(3) = &H5E
    testBytes(4) = &HFF
    
    ' standard base64 for these bytes: "/D+QXv8=" (using + and / and padding)
    ' base64url expected: "_D-QXv8"
    Dim result As String
    result = CryptoUtils.Base64URLEncode(testBytes)
    
    If result <> "_D-QXv8" Then
        Debug.Print "TestBase64URL Fail: Expected '_D-QXv8' but got '" & result & "'"
        TestBase64URL = False
        Exit Function
    End If
    
    Debug.Print "TestBase64URL: PASSED"
    TestBase64URL = True
    Exit Function

ErrorHandler:
    Debug.Print "TestBase64URL Error: " & Err.Description
    TestBase64URL = False
End Function

' ==============================================================================
' Test Code Verifier and Challenge Generation
' ==============================================================================
Private Function TestCodeVerifierAndChallenge() As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Test Code Verifier constraints
    Dim verifier As String
    verifier = CryptoUtils.GenerateCodeVerifier(80)
    
    If Len(verifier) <> 80 Then
        Debug.Print "TestCodeVerifierAndChallenge Fail: Expected length 80, got " & Len(verifier)
        TestCodeVerifierAndChallenge = False
        Exit Function
    End If
    
    ' Ensure characters are only in unreserved set: [A-Za-z0-9-._~]
    Dim i As Long
    Dim char As String
    Dim allowed As String
    allowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
    
    For i = 1 To Len(verifier)
        char = Mid$(verifier, i, 1)
        If InStr(1, allowed, char) = 0 Then
            Debug.Print "TestCodeVerifierAndChallenge Fail: Invalid character '" & char & "' found in verifier."
            TestCodeVerifierAndChallenge = False
            Exit Function
        End If
    Next i
    
    ' 2. Verify that two verifiers generated consecutively are different (entropy check)
    Dim verifier2 As String
    verifier2 = CryptoUtils.GenerateCodeVerifier(80)
    If verifier = verifier2 Then
        Debug.Print "TestCodeVerifierAndChallenge Fail: Generated identical verifiers."
        TestCodeVerifierAndChallenge = False
        Exit Function
    End If
    
    ' 3. Verify Code Challenge generation works without error
    Dim challenge As String
    challenge = CryptoUtils.GenerateCodeChallenge(verifier)
    If Len(challenge) < 43 Then
        Debug.Print "TestCodeVerifierAndChallenge Fail: Challenge too short: " & challenge
        TestCodeVerifierAndChallenge = False
        Exit Function
    End If
    
    Debug.Print "TestCodeVerifierAndChallenge: PASSED"
    TestCodeVerifierAndChallenge = True
    Exit Function

ErrorHandler:
    Debug.Print "TestCodeVerifierAndChallenge Error: " & Err.Description
    TestCodeVerifierAndChallenge = False
End Function

' ==============================================================================
' Test JSON Parser
' ==============================================================================
Private Function TestJSONParser() As Boolean
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = "{""access_token"":""TOKEN123"",""instance_url"":""https://salesforce.com"",""is_active"":true,""total_records"":45,""null_val"":null,""nested"":{""key"":""val""},""list"":[1,""two"",3]}"
    
    Dim parsed As Object
    Set parsed = JSONParser.ParseJSON(json)
    
    ' Validate flat values
    If parsed("access_token") <> "TOKEN123" Then
        Debug.Print "TestJSONParser Fail: access_token mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    If parsed("instance_url") <> "https://salesforce.com" Then
        Debug.Print "TestJSONParser Fail: instance_url mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    If parsed("is_active") <> True Then
        Debug.Print "TestJSONParser Fail: is_active mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    If parsed("total_records") <> 45 Then
        Debug.Print "TestJSONParser Fail: total_records mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    If Not IsNull(parsed("null_val")) Then
        Debug.Print "TestJSONParser Fail: null_val mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    ' Validate nested object
    Dim nested As Object
    Set nested = parsed("nested")
    If nested("key") <> "val" Then
        Debug.Print "TestJSONParser Fail: nested key mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    ' Validate list array
    Dim list As Object
    Set list = parsed("list")
    If list.Count <> 3 Then
        Debug.Print "TestJSONParser Fail: list count mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    If list(1) <> 1 Or list(2) <> "two" Or list(3) <> 3 Then
        Debug.Print "TestJSONParser Fail: list elements mismatch"
        TestJSONParser = False
        Exit Function
    End If
    
    Debug.Print "TestJSONParser: PASSED"
    TestJSONParser = True
    Exit Function

ErrorHandler:
    Debug.Print "TestJSONParser Error: " & Err.Description
    TestJSONParser = False
End Function
