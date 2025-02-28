Function CallOpenAIAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String
    
    ' OpenAI API endpoint for chat completions (adjust to your use case)
    API = "https://api.openai.com/v1/chat/completions"
    
    ' Prepare request payload for OpenAI
    ' SendTxt = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": ""You are a Word assistant""}, {""role"": ""user"", ""content"": """ & inputText & """}], ""max_tokens"": 4096}"
    SendTxt = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": ""You are a Word assistant""}, {""role"": ""user"", ""content"": """ & inputText & """}], ""max_tokens"": 4096}"
    
    ' Create HTTP object
    Set Http = CreateObject("MSXML2.XMLHTTP")
    
    ' Open connection to OpenAI API
    With Http
        .Open "POST", API, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & api_key
        .send SendTxt
        status_code = .Status
        response = .responseText
    End With
    
    ' Display API response (for debugging purposes)
    MsgBox "API Response: " & response, vbInformation, "Debug Info"
    
    If status_code = 200 Then
        CallOpenAIAPI = response
    Else
        CallOpenAIAPI = "Error: " & status_code & " - " & response
    End If
    
    Set Http = Nothing
End Function

Sub OpenAI()
    Dim api_key As String
    Dim inputText As String
    Dim response As String
    Dim regex As Object
    Dim contentRegex As Object
    Dim matches As Object
    Dim originalSelection As Object
    Dim finalContent As String
    
    api_key = "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" ' Replace with your OpenAI API key
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If
    
    ' Save the original selection text
    Set originalSelection = Selection.Range.Duplicate
    
    ' Clean and prepare the input text for API call
    inputText = Replace(Replace(Replace(Replace(Replace(Selection.text, "\", "\\"), vbCrLf, ""), vbCr, ""), vbLf, ""), Chr(34), "\""")
    
    ' Call the OpenAI API with the input text
    response = CallOpenAIAPI(api_key, inputText)
    
    If Left(response, 5) <> "Error" Then
        ' Create regular expression objects to match the response content
        Set contentRegex = CreateObject("VBScript.RegExp")
        With contentRegex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = """content"":\s*""([^""\\]*(\\.[^""\\]*)*)"""
        End With
        
        ' Extract the final response content
        Set matches = contentRegex.Execute(response)
        If matches.Count > 0 Then
            finalContent = matches(0).SubMatches(0)
            finalContent = Replace(finalContent, "\n\n", vbNewLine)
            finalContent = Replace(finalContent, "\n", vbNewLine)
            finalContent = Replace(Replace(finalContent, """", Chr(34)), """", Chr(34))
            
            ' Insert the final response into the Word document
            Selection.TypeText finalContent
            ' Selection.TypeText response
            Selection.TypeParagraph  ' Add an extra paragraph break
            Selection.Collapse Direction:=wdCollapseEnd  ' Move cursor to the end
        Else
            MsgBox "Failed to parse API response.", vbExclamation
        End If
    Else
        MsgBox response, vbCritical
    End If
End Sub

