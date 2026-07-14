Attribute VB_Name = "CryptoUtils"
Option Explicit

' ==============================================================================
' CryptoUtils.bas - Windows CNG Cryptographic Helpers for VBA
' Compatible with 32-bit and 64-bit Excel/Office (VBA7)
' Handles secure random generation, UTF-8 conversion, SHA256, and Base64URL
' ==============================================================================

' WideCharToMultiByte API for Unicode to UTF-8 conversion
#If VBA7 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private Const CP_UTF8 As Long = 65001

' Windows CNG (Cryptography Next Generation) BCrypt APIs
#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal pbBuffer As LongPtr, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal pbBuffer As Long, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
#End If

' BCrypt API Constants
Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2
Private Const STATUS_SUCCESS As Long = &H0

' ==============================================================================
' Cryptographically Secure Random Byte Generator
' ==============================================================================
Public Function GetRandomBytes(ByVal Length As Long) As Byte()
    Dim buffer() As Byte
    ReDim buffer(0 To Length - 1)
    
    Dim status As Long
    status = BCryptGenRandom(0, VarPtr(buffer(0)), Length, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    
    If status <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1001, "CryptoUtils", "Failed to generate random bytes using BCrypt. Status: " & status
    End If
    
    GetRandomBytes = buffer
End Function

' ==============================================================================
' PKCE Code Verifier Generator
' Generates an unbiased random string of size 43 to 128 containing [A-Za-z0-9-._~]
' ==============================================================================
Public Function GenerateCodeVerifier(Optional ByVal Length As Long = 64) As String
    If Length < 43 Or Length > 128 Then
        Err.Raise vbObjectError + 1002, "CryptoUtils", "Code verifier length must be between 43 and 128 characters."
    End If
    
    Dim chars As String
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
    
    Dim result As String
    result = Space$(Length)
    
    Dim randomByte As Byte
    Dim i As Long
    i = 1
    
    ' Loop until we get the requested length of characters
    Do While i <= Length
        ' Get a single cryptographically secure random byte
        randomByte = GetRandomBytes(1)(0)
        
        ' 3 * 66 = 198. To prevent modulo bias, discard values >= 198
        If randomByte < 198 Then
            Mid$(result, i, 1) = Mid$(chars, (randomByte Mod 66) + 1, 1)
            i = i + 1
        End If
    Loop
    
    GenerateCodeVerifier = result
End Function

' ==============================================================================
' Convert String to UTF-8 Byte Array
' ==============================================================================
Public Function ConvertToUTF8(ByVal InputStr As String) As Byte()
    Dim buffer() As Byte
    If Len(InputStr) = 0 Then
        ConvertToUTF8 = buffer
        Exit Function
    End If
    
    Dim utf8Len As Long
    utf8Len = WideCharToMultiByte(CP_UTF8, 0, StrPtr(InputStr), Len(InputStr), 0, 0, 0, 0)
    
    ReDim buffer(0 To utf8Len - 1)
    WideCharToMultiByte(CP_UTF8, 0, StrPtr(InputStr), Len(InputStr), VarPtr(buffer(0)), utf8Len, 0, 0)
    
    ConvertToUTF8 = buffer
End Function

' ==============================================================================
' SHA-256 Hashing of a String using Windows CNG
' ==============================================================================
Public Function GetSHA256Hash(ByVal InputStr As String) As Byte()
#If VBA7 Then
    Dim hAlg As LongPtr
    Dim hHash As LongPtr
#Else
    Dim hAlg As Long
    Dim hHash As Long
#End If
    
    Dim status As Long
    status = BCryptOpenAlgorithmProvider(hAlg, StrPtr(BCRYPT_SHA256_ALGORITHM), 0, 0)
    If status <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1003, "CryptoUtils", "Failed to open algorithm provider. Status: " & status
    End If
    
    ' Query object and digest sizes
    Dim hashObjectSize As Long
    Dim cbResult As Long
    status = BCryptGetProperty(hAlg, StrPtr(BCRYPT_OBJECT_LENGTH), VarPtr(hashObjectSize), 4, cbResult, 0)
    If status <> STATUS_SUCCESS Then
        BCryptCloseAlgorithmProvider hAlg, 0
        Err.Raise vbObjectError + 1004, "CryptoUtils", "Failed to get hash object size. Status: " & status
    End If
    
    Dim hashSize As Long
    status = BCryptGetProperty(hAlg, StrPtr(BCRYPT_HASH_LENGTH), VarPtr(hashSize), 4, cbResult, 0)
    If status <> STATUS_SUCCESS Then
        BCryptCloseAlgorithmProvider hAlg, 0
        Err.Raise vbObjectError + 1005, "CryptoUtils", "Failed to get hash digest size. Status: " & status
    End If
    
    ' Allocate hash object memory buffer
    Dim hashObject() As Byte
    ReDim hashObject(0 To hashObjectSize - 1)
    
    ' Create the hash object
    status = BCryptCreateHash(hAlg, hHash, VarPtr(hashObject(0)), hashObjectSize, 0, 0, 0)
    If status <> STATUS_SUCCESS Then
        BCryptCloseAlgorithmProvider hAlg, 0
        Err.Raise vbObjectError + 1006, "CryptoUtils", "Failed to create hash object. Status: " & status
    End If
    
    ' Convert input to UTF-8
    Dim inputBytes() As Byte
    inputBytes = ConvertToUTF8(InputStr)
    
    ' Hash data
    If UBound(inputBytes) >= 0 Then
        status = BCryptHashData(hHash, VarPtr(inputBytes(0)), UBound(inputBytes) + 1, 0)
        If status <> STATUS_SUCCESS Then
            BCryptDestroyHash hHash
            BCryptCloseAlgorithmProvider hAlg, 0
            Err.Raise vbObjectError + 1007, "CryptoUtils", "Failed to hash data. Status: " & status
        End If
    End If
    
    ' Finalize hash to get digest bytes
    Dim hashResult() As Byte
    ReDim hashResult(0 To hashSize - 1)
    status = BCryptFinishHash(hHash, VarPtr(hashResult(0)), hashSize, 0)
    If status <> STATUS_SUCCESS Then
        BCryptDestroyHash hHash
        BCryptCloseAlgorithmProvider hAlg, 0
        Err.Raise vbObjectError + 1008, "CryptoUtils", "Failed to finish hash. Status: " & status
    End If
    
    ' Cleanup resources
    BCryptDestroyHash hHash
    BCryptCloseAlgorithmProvider hAlg, 0
    
    GetSHA256Hash = hashResult
End Function

' ==============================================================================
' Standard Base64 Encoding using MSXML2.DOMDocument
' ==============================================================================
Public Function Base64Encode(ByRef Bytes() As Byte) As String
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    Dim xmlElem As Object
    Set xmlElem = xmlDoc.createElement("tmp")
    xmlElem.DataType = "bin.base64"
    xmlElem.NodeTypedValue = Bytes
    
    Dim base64Str As String
    base64Str = xmlElem.Text
    
    ' Remove line wraps formatting from XML output
    base64Str = Replace(base64Str, vbLf, "")
    base64Str = Replace(base64Str, vbCr, "")
    
    Base64Encode = base64Str
End Function

' ==============================================================================
' RFC 4648 Base64URL Encoding (URL-safe without padding)
' ==============================================================================
Public Function Base64URLEncode(ByRef Bytes() As Byte) As String
    Dim base64 As String
    base64 = Base64Encode(Bytes)
    
    ' Replace characters for URL-safety
    base64 = Replace(base64, "+", "-")
    base64 = Replace(base64, "/", "_")
    
    ' Strip trailing padding characters
    Do While Right$(base64, 1) = "="
        base64 = Left$(base64, Len(base64) - 1)
    Loop
    
    Base64URLEncode = base64
End Function

' ==============================================================================
' PKCE Code Challenge Generator
' ==============================================================================
Public Function GenerateCodeChallenge(ByVal CodeVerifier As String) As String
    Dim hash() As Byte
    hash = GetSHA256Hash(CodeVerifier)
    GenerateCodeChallenge = Base64URLEncode(hash)
End Function
