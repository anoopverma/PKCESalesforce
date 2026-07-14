Attribute VB_Name = "JSONParser"
Option Explicit

' ==============================================================================
' JSONParser.bas - Lightweight Recursive-Descent JSON Parser for VBA
' Parses JSON strings into Scripting.Dictionary (Objects) and Collections (Arrays).
' No external ActiveX references needed (except Microsoft Scripting Runtime).
' ==============================================================================

Private Index As Long
Private JSONText As String
Private Length As Long

' ==============================================================================
' Main Entry Point
' ==============================================================================
Public Function ParseJSON(ByVal JSONString As String) As Object
    JSONText = JSONString
    Length = Len(JSONText)
    Index = 1
    SkipWhitespace
    
    If Index > Length Then
        Err.Raise vbObjectError + 2001, "JSONParser", "JSON string is empty."
    End If
    
    Dim char As String
    char = Mid$(JSONText, Index, 1)
    
    If char = "{" Then
        Set ParseJSON = ParseObject()
    ElseIf char = "[" Then
        Set ParseJSON = ParseArray()
    Else
        Err.Raise vbObjectError + 2002, "JSONParser", "JSON must start with '{' or '[' at index " & Index
    End If
End Function

' ==============================================================================
' Helper to Assign Object or Value to Variant
' ==============================================================================
Private Sub AssignValue(ByRef Target As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Target = Value
    Else
        Target = Value
    End If
End Sub

' ==============================================================================
' Skip Whitespace Characters
' ==============================================================================
Private Sub SkipWhitespace()
    Dim char As String
    Do While Index <= Length
        char = Mid$(JSONText, Index, 1)
        If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
            Index = Index + 1
        Else
            Exit Sub
        End If
    Loop
End Sub

' ==============================================================================
' Value Parser (Orchestrator)
' ==============================================================================
Private Function ParseValue() As Variant
    SkipWhitespace
    If Index > Length Then
        Err.Raise vbObjectError + 2003, "JSONParser", "Unexpected end of JSON input."
    End If
    
    Dim char As String
    char = Mid$(JSONText, Index, 1)
    
    If char = """" Then
        ParseValue = ParseString()
    ElseIf char = "{" Then
        Set ParseValue = ParseObject()
    ElseIf char = "[" Then
        Set ParseValue = ParseArray()
    ElseIf char = "t" Or char = "f" Then
        ParseValue = ParseBoolean()
    ElseIf char = "n" Then
        ParseValue = ParseNull()
    ElseIf char = "-" Or (char >= "0" And char <= "9") Then
        ParseValue = ParseNumber()
    Else
        Err.Raise vbObjectError + 2004, "JSONParser", "Unexpected character '" & char & "' at index " & Index
    End If
End Function

' ==============================================================================
' Object Parser (maps to Scripting.Dictionary)
' ==============================================================================
Private Function ParseObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' Case-insensitive keys
    
    ' Skip '{'
    Index = Index + 1
    SkipWhitespace
    
    If Mid$(JSONText, Index, 1) = "}" Then
        Index = Index + 1
        Set ParseObject = dict
        Exit Function
    End If
    
    Do
        SkipWhitespace
        If Mid$(JSONText, Index, 1) <> """" Then
            Err.Raise vbObjectError + 2005, "JSONParser", "Expected key string at index " & Index
        End If
        
        Dim key As String
        key = ParseString()
        
        SkipWhitespace
        If Mid$(JSONText, Index, 1) <> ":" Then
            Err.Raise vbObjectError + 2006, "JSONParser", "Expected ':' separator at index " & Index
        End If
        Index = Index + 1 ' Skip ':'
        
        Dim val As Variant
        AssignValue val, ParseValue()
        
        If IsObject(val) Then
            Set dict(key) = val
        Else
            dict(key) = val
        End If
        
        SkipWhitespace
        Dim char As String
        char = Mid$(JSONText, Index, 1)
        If char = "," Then
            Index = Index + 1 ' Skip ','
        ElseIf char = "}" Then
            Index = Index + 1 ' Skip '}'
            Exit Do
        Else
            Err.Raise vbObjectError + 2007, "JSONParser", "Expected ',' or '}' at index " & Index
        End If
    Loop
    
    Set ParseObject = dict
End Function

' ==============================================================================
' Array Parser (maps to Collection)
' ==============================================================================
Private Function ParseArray() As Object
    Dim col As Object
    Set col = New Collection
    
    ' Skip '['
    Index = Index + 1
    SkipWhitespace
    
    If Mid$(JSONText, Index, 1) = "]" Then
        Index = Index + 1
        Set ParseArray = col
        Exit Function
    End If
    
    Do
        Dim val As Variant
        AssignValue val, ParseValue()
        
        col.Add val
        
        SkipWhitespace
        Dim char As String
        char = Mid$(JSONText, Index, 1)
        If char = "," Then
            Index = Index + 1 ' Skip ','
        ElseIf char = "]" Then
            Index = Index + 1 ' Skip ']'
            Exit Do
        Else
            Err.Raise vbObjectError + 2008, "JSONParser", "Expected ',' or ']' at index " & Index
        End If
    Loop
    
    Set ParseArray = col
End Function

' ==============================================================================
' String Parser (supports escape sequences and unicode escapes)
' ==============================================================================
Private Function ParseString() As String
    ' Skip opening quote
    Index = Index + 1
    
    Dim result As String
    result = ""
    Dim char As String
    
    Do While Index <= Length
        char = Mid$(JSONText, Index, 1)
        
        If char = """" Then
            Index = Index + 1 ' Skip closing quote
            ParseString = result
            Exit Function
        ElseIf char = "\" Then
            Index = Index + 1
            If Index > Length Then
                Err.Raise vbObjectError + 2009, "JSONParser", "Unterminated escape sequence in JSON string."
            End If
            
            Dim escapeChar As String
            escapeChar = Mid$(JSONText, Index, 1)
            Index = Index + 1
            
            Select Case escapeChar
                Case """"
                    result = result & """"
                Case "\"
                    result = result & "\"
                Case "/"
                    result = result & "/"
                Case "b"
                    result = result & Chr(8)  ' Backspace
                Case "f"
                    result = result & Chr(12) ' Form feed
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    ' 4-hex digit unicode escape \uXXXX
                    If Index + 3 > Length Then
                        Err.Raise vbObjectError + 2010, "JSONParser", "Invalid Unicode escape sequence."
                    End If
                    Dim hexCode As String
                    hexCode = Mid$(JSONText, Index, 4)
                    Index = Index + 4
                    result = result & ChrW(Val("&H" & hexCode))
                Case Else
                    result = result & escapeChar
            End Select
        Else
            result = result & char
            Index = Index + 1
        End If
    Loop
    
    Err.Raise vbObjectError + 2011, "JSONParser", "Unterminated string in JSON input."
End Function

' ==============================================================================
' Boolean Parser
' ==============================================================================
Private Function ParseBoolean() As Boolean
    If Mid$(JSONText, Index, 4) = "true" Then
        Index = Index + 4
        ParseBoolean = True
    ElseIf Mid$(JSONText, Index, 5) = "false" Then
        Index = Index + 5
        ParseBoolean = False
    Else
        Err.Raise vbObjectError + 2012, "JSONParser", "Invalid boolean value."
    End If
End Function

' ==============================================================================
' Null Parser
' ==============================================================================
Private Function ParseNull() As Variant
    If Mid$(JSONText, Index, 4) = "null" Then
        Index = Index + 4
        ParseNull = Null
    Else
        Err.Raise vbObjectError + 2013, "JSONParser", "Invalid null value."
    End If
End Function

' ==============================================================================
' Number Parser
' ==============================================================================
Private Function ParseNumber() As Double
    Dim start As Long
    start = Index
    
    Dim char As String
    Do While Index <= Length
        char = Mid$(JSONText, Index, 1)
        If (char >= "0" And char <= "9") Or char = "." Or char = "-" Or char = "+" Or char = "e" Or char = "E" Then
            Index = Index + 1
        Else
            Exit Do
        End If
    Loop
    
    Dim numStr As String
    numStr = Mid$(JSONText, start, Index - start)
    
    ' Val is locale-independent and parses period decimals correctly
    ParseNumber = Val(numStr)
End Function
